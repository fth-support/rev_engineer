import React from 'react'

function Journey() {
  return (
    <div className="journey-page">
      <h1>Application Journey</h1>
      <p>ภาพรวมการเดินทางของข้อมูล (Data Flow) ตั้งแต่ลูกค้าซื้อสินค้าหน้าสาขา ไปจนถึงข้อมูลไปรวมที่สำนักงานใหญ่</p>
      
      <div style={{ marginTop: '2rem' }}>
        <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1.5rem', borderLeft: '4px solid #3b82f6' }}>
          <h3>1. Transaction Capture (Offline First)</h3>
          <p>เมื่อลูกค้าซื้อสินค้าที่เครื่อง POS ข้อมูลจะถูกบันทึกลงฐานข้อมูล Local (SQL Express) ทันที โดยมีสถานะเริ่มต้นคือรอการซิงค์ (FTStaSentOnOff = '0')</p>
        </div>

        <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1.5rem', borderLeft: '4px solid #8b5cf6' }}>
          <h3>2. Background Synchronization</h3>
          <p>ServiceTransfer ทำงานแบบ Polling Loop ทุกๆ 500ms ตรวจสอบบิลใหม่เพื่อเตรียมจัดส่ง</p>
        </div>

        <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1.5rem', borderLeft: '4px solid #f59e0b' }}>
          <h3>3. Data Tokenization (Security)</h3>
          <p>เลขบัตรเครดิตและเลขบัตรสมาชิก จะถูกส่งไปสอบถาม Token ผ่านทาง SafeNet Web Service ระบบจะเก็บเฉพาะ Token เพื่อส่งต่อขึ้น HQ</p>
        </div>

        <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1.5rem', borderLeft: '4px solid #10b981' }}>
          <h3>4. Central Delivery</h3>
          <p>บันทึกข้อมูลธุรกรรมที่ถูก Tokenize แล้ว ลงสู่ Central DB ที่สำนักงานใหญ่ พร้อมคำนวณและอัปเดตคะแนนสมาชิกในระบบ MallCard</p>
        </div>
      </div>
    </div>
  )
}

export default Journey
