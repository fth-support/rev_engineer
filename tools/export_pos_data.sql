/* =============================================================================
   export_pos_data.sql   —   ONE SCRIPT TO RULE THEM ALL
   ส่งออก schema + ข้อมูลทุกตารางของ POSSDB เป็นไฟล์ CSV ในคลิกเดียว
   -----------------------------------------------------------------------------
   วิธีใช้ (ง่ายมาก 3 ขั้น):
     1) เปิด SSMS แล้ว connect เข้า SQL Server ของ POS  *เป็น admin หรือ sa*
        (จำเป็น เพราะสคริปต์ต้องเปิด xp_cmdshell ชั่วคราวเพื่อเขียนไฟล์)
     2) เปิดไฟล์นี้ → ถ้าชื่อ DB ไม่ใช่ POSSDB ให้แก้บรรทัด USE ข้างล่าง
     3) กด F5 (Execute) ทีเดียว → รอจนขึ้น  *** DONE ***  ใน tab Messages
   ผลลัพธ์: ไฟล์ CSV ทุกตารางอยู่ใน  C:\pos_export\
     - ตารางเล็ก (<= 5000 แถว)  = เอาครบทุกแถว
     - ตารางใหญ่                 = เอา 300 แถวล่าสุดพอ
     - ไฟล์สรุป: _columns.csv (schema เต็ม) · _keys.csv · _row_counts.csv
   จากนั้น: ก๊อปโฟลเดอร์  C:\pos_export\  ทั้งโฟลเดอร์ ใส่ในโฟลเดอร์ repo
   (ตั้งชื่อ db_export) แล้วบอก Claude — Claude อ่าน local; ไม่ขึ้น git
   -----------------------------------------------------------------------------
   หมายเหตุ
   • ตัวคั่นในไฟล์คือ  |@|  (กันชนกับ comma ในข้อมูล) — Claude อ่านได้ ไม่ต้องห่วง
   • เป็นข้อมูลจริง (มี PII) → เก็บไว้ใน db_export\ ที่ git-ignore ไว้แล้ว อย่า push เอง
   • ถ้ากด F5 แล้วขึ้น error เรื่องสิทธิ์ = ยังไม่ได้ connect เป็น admin/sa
     → ปิด SSMS, คลิกขวา "Run as administrator", connect ใหม่ แล้วลองอีกครั้ง
   ============================================================================= */

USE POSSDB;     -- << แก้ตรงนี้ถ้าชื่อฐานข้อมูลไม่ใช่ POSSDB
GO

SET NOCOUNT ON;

DECLARE @OutDir    nvarchar(260) = N'C:\pos_export';  -- โฟลเดอร์ปลายทาง (เปลี่ยนได้)
DECLARE @BigSample int = 300;                          -- ตารางใหญ่เอากี่แถว
DECLARE @SmallMax  int = 5000;                         -- <= เท่านี้ = เอาครบทุกแถว

/* 1) เปิด xp_cmdshell ชั่วคราว ----------------------------------------------- */
BEGIN TRY
    EXEC sp_configure 'show advanced options', 1; RECONFIGURE;
    EXEC sp_configure 'xp_cmdshell', 1; RECONFIGURE;
END TRY
BEGIN CATCH
    PRINT '!! เปิด xp_cmdshell ไม่ได้ — ต้อง connect เป็น admin/sa';
    PRINT '   ปิด SSMS แล้วเปิดใหม่แบบ Run as administrator, connect ใหม่ แล้วรันอีกครั้ง';
    RETURN;
END CATCH

DECLARE @srv sysname = @@SERVERNAME;
DECLARE @db  sysname = DB_NAME();
DECLARE @cmd nvarchar(4000);

/* 2) สร้างโฟลเดอร์ปลายทาง ---------------------------------------------------- */
SET @cmd = N'if not exist "' + @OutDir + N'" mkdir "' + @OutDir + N'"';
EXEC master..xp_cmdshell @cmd, no_output;

/* 3) ไฟล์สรุป schema (3 ไฟล์) ------------------------------------------------ */
-- _columns.csv : ทุกคอลัมน์ + ชนิดข้อมูล
SET @cmd = N'bcp "SET NOCOUNT ON; SELECT c.TABLE_NAME,c.ORDINAL_POSITION,c.COLUMN_NAME,c.DATA_TYPE,'
         + N'c.CHARACTER_MAXIMUM_LENGTH,c.NUMERIC_PRECISION,c.NUMERIC_SCALE,c.IS_NULLABLE '
         + N'FROM [' + @db + N'].INFORMATION_SCHEMA.COLUMNS c '
         + N'JOIN [' + @db + N'].INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME=c.TABLE_NAME AND t.TABLE_TYPE=''BASE TABLE'' '
         + N'ORDER BY c.TABLE_NAME,c.ORDINAL_POSITION" '
         + N'queryout "' + @OutDir + N'\_columns.csv" -c -t"|@|" -T -S ' + @srv;
EXEC master..xp_cmdshell @cmd, no_output;

-- _keys.csv : primary / foreign keys
SET @cmd = N'bcp "SET NOCOUNT ON; SELECT tc.CONSTRAINT_TYPE,kcu.TABLE_NAME,kcu.COLUMN_NAME,kcu.ORDINAL_POSITION,kcu.CONSTRAINT_NAME '
         + N'FROM [' + @db + N'].INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc '
         + N'JOIN [' + @db + N'].INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON tc.CONSTRAINT_NAME=kcu.CONSTRAINT_NAME '
         + N'WHERE tc.CONSTRAINT_TYPE IN (''PRIMARY KEY'',''FOREIGN KEY'') '
         + N'ORDER BY kcu.TABLE_NAME,tc.CONSTRAINT_TYPE,kcu.ORDINAL_POSITION" '
         + N'queryout "' + @OutDir + N'\_keys.csv" -c -t"|@|" -T -S ' + @srv;
EXEC master..xp_cmdshell @cmd, no_output;

-- _row_counts.csv : จำนวนแถวทุกตาราง
SET @cmd = N'bcp "SET NOCOUNT ON; SELECT t.name,SUM(p.rows) AS rws '
         + N'FROM [' + @db + N'].sys.tables t JOIN [' + @db + N'].sys.partitions p '
         + N'ON p.object_id=t.object_id AND p.index_id IN (0,1) GROUP BY t.name ORDER BY rws DESC" '
         + N'queryout "' + @OutDir + N'\_row_counts.csv" -c -t"|@|" -T -S ' + @srv;
EXEC master..xp_cmdshell @cmd, no_output;

/* 4) วนทุกตาราง → เซฟ CSV (เล็กเอาครบ / ใหญ่เอา 300) ------------------------- */
DECLARE @t sysname, @rows bigint, @top nvarchar(20);
DECLARE cur CURSOR FAST_FORWARD FOR
    SELECT t.name, SUM(p.rows)
    FROM sys.tables t
    JOIN sys.partitions p ON p.object_id = t.object_id AND p.index_id IN (0,1)
    GROUP BY t.name;
OPEN cur; FETCH NEXT FROM cur INTO @t, @rows;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @top = CASE WHEN @rows <= @SmallMax THEN N''
                    ELSE N'TOP ' + CAST(@BigSample AS nvarchar(10)) + N' ' END;
    SET @cmd = N'bcp "SET NOCOUNT ON; SELECT ' + @top + N'* FROM [' + @db + N'].[dbo].[' + @t + N']" '
             + N'queryout "' + @OutDir + N'\' + @t + N'.csv" -c -t"|@|" -T -S ' + @srv;
    EXEC master..xp_cmdshell @cmd, no_output;
    FETCH NEXT FROM cur INTO @t, @rows;
END
CLOSE cur; DEALLOCATE cur;

/* 5) ปิด xp_cmdshell คืน (ความปลอดภัย) -------------------------------------- */
EXEC sp_configure 'xp_cmdshell', 0; RECONFIGURE;
EXEC sp_configure 'show advanced options', 0; RECONFIGURE;

/* 6) โชว์รายชื่อไฟล์ที่ได้ ให้เห็นว่าเสร็จจริง ------------------------------- */
SET @cmd = N'dir "' + @OutDir + N'\*.csv"';
EXEC master..xp_cmdshell @cmd;

PRINT '';
PRINT '*** DONE *** ไฟล์ทั้งหมดอยู่ที่ ' + @OutDir;
PRINT 'ก๊อปโฟลเดอร์นี้ใส่ในโฟลเดอร์ repo (ตั้งชื่อ db_export) แล้วบอก Claude ได้เลย';
GO
