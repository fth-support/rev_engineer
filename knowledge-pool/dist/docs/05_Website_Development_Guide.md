# 🚀 Website Development Guide: ServiceTransfer Documentation Portal

เอกสารฉบับนี้สรุปโครงสร้าง สถาปัตยกรรม และเทคนิคที่ใช้ในการสร้างเว็บไซต์ **ServiceTransfer Documentation Portal** นี้ขึ้นมา เพื่อเป็นแม่แบบ (Template) สำหรับนำไปประยุกต์ใช้หรือต่อยอดในโปรเจกต์อื่นๆ ต่อไป

---

## 🛠️ 1. เทคโนโลยีหลัก (Tech Stack)
โปรเจกต์นี้ถูกออกแบบมาให้ทำงานได้รวดเร็ว เบา และไม่ต้องพึ่งพาไลบรารีที่หนักเกินความจำเป็น:
- **Core Framework:** React 18 + Vite (รวดเร็วในการ Build และ Hot Reload)
- **Routing:** `react-router-dom` (จัดการหน้าต่างๆ แบบ Single Page Application - SPA)
- **Styling:** Pure CSS (ใช้ CSS Variables, CSS Grid และ Flexbox) ไม่ใช้ Tailwind เพื่อความเข้าใจง่ายและการปรับแต่งระดับละเอียด
- **Markdown Rendering:** `react-markdown` + `remark-gfm`
- **Icons:** `lucide-react` (เบา สวยงาม โหลดเฉพาะไอคอนที่ใช้)

---

## 📐 2. โครงสร้าง Layout สไตล์ "Bento Grid"
เว็บนี้ใช้ดีไซน์ที่กำลังเป็นที่นิยมคือ **Bento Grid** (การจัดเรียงกล่องสี่เหลี่ยมคล้ายกล่องข้าวเบนโตะ) เพื่อการนำเสนอข้อมูลหลายๆ ส่วนในหน้าเดียวอย่างเป็นระเบียบ

### เทคนิคที่ใช้ (CSS Grid Areas)
ในไฟล์ `src/Bento.css` มีการกำหนด `grid-template-areas` ซึ่งทำให้การจัดวางกล่อง (Cells) ทำได้ง่ายมากในแต่ละขนาดหน้าจอ:

```css
/* สำหรับหน้าจอ Desktop */
@media (min-width: 1024px) {
  .bento {
      grid-template-columns: repeat(4, 1fr);
      grid-template-areas:
          'hero  hero  about stats'
          'media feed  tasks card'
          'promo promo promo promo'
          'feat  feat  feat  feat';
  }
}
```
**การประยุกต์ใช้:** หากคุณต้องการสลับตำแหน่งกล่องในอนาคต เพียงแค่สลับชื่อใน `grid-template-areas` โครงสร้างเว็บก็จะเปลี่ยนตามทันทีโดยไม่ต้องแก้ HTML เลย!

---

## 📄 3. ระบบแสดงผลเอกสาร Markdown อัจฉริยะ (Smart Markdown Viewer)
หัวใจหลักของเว็บนี้คือการดึงไฟล์ `.md` จากโฟลเดอร์ `public/docs/` มาแสดงผล ซึ่งจัดการโดยคอมโพเนนต์ `MarkdownViewer.jsx`

### ความลับของการแทรก React Components ลงใน Markdown
โดยปกติ Markdown จะพิมพ์ได้แค่ข้อความ แต่เว็บนี้สามารถ **"แทรกแผนภาพ React แบบ Interactive เข้าไปกลางหน้ากระดาษได้เลย"** 

**วิธีทำ:**
1. ในไฟล์ Markdown เราตั้งชื่อภาษาโค้ดแบบกำหนดเอง (Custom Language) เช่น:
   ````text
   ```diagram
   architecture
   ```
   ````
2. ใน `MarkdownViewer.jsx` เราเข้าไปแทรกแซงกระบวนการเรนเดอร์ของ `<ReactMarkdown>`:
   ```jsx
   components={{
     code({node, inline, className, children, ...props}) {
       const match = /language-(\w+)/.exec(className || '');
       // หากพบคำว่า 'diagram' จะเรียกใช้ DiagramRenderer ทันที
       if (!inline && match && match[1] === 'diagram') {
         return <DiagramRenderer id={children} />
       }
       return <code className={className} {...props}>{children}</code>;
     }
   }}
   ```
**การประยุกต์ใช้:** คุณสามารถใช้เทคนิคนี้เพื่อแทรกอะไรก็ได้ลงในบทความ เช่น กราฟ สถิติ วิดีโอ หรือแม้กระทั่งมินิเกม!

---

## 🎨 4. Native React Diagrams (แทนที่ภาพนิ่ง/Iframe/Mermaid)
เพื่อให้เว็บดูพรีเมียม (Premium) แผนภาพต่างๆ (Flows/Architecture) จึงถูกเขียนขึ้นมาเป็น Native React Components ล้วนๆ แทนที่จะใช้ภาพนิ่ง หรือไลบรารีที่หนักอย่าง Mermaid

- โค้ดแผนภาพอยู่ในโฟลเดอร์ `src/components/diagrams/`
- จัดรูปแบบและสีสันด้วยคลาสในไฟล์ `Diagrams.css` 
- การเขียนแบบนี้ทำให้กล่องข้อมูล หรือลูกศร สามารถตอบสนอง (Responsive) ต่อขนาดหน้าจอได้ และสามารถเพิ่มแอนิเมชันตอนชี้เมาส์ (Hover) ในอนาคตได้ง่าย

---

## 🔗 5. ระบบ Routing และ Menu Highlight สัมพันธ์กัน (Synchronization)
เวลาที่เรากดกล่องเอกสารที่หน้า `Home.jsx` เมนูด้านซ้าย (Sidebar) ก็จะสว่าง (Active) ตามโดยอัตโนมัติ

**เบื้องหลังการทำงาน:**
ไม่ได้ใช้ State ที่ซับซ้อน แต่พึ่งพากลไกของ `<NavLink>` ใน React Router:
```jsx
// ใน Sidebar.jsx
<NavLink 
  to="/docs/architecture-srs" 
  className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`}
>
```
เมื่อหน้าจอใดๆ สั่งเปลี่ยน URL (เช่น ไปที่ `/docs/architecture-srs`) `NavLink` ที่มี `to` ตรงกับ URL นั้น จะได้รับสถานะ `isActive = true` ทันที ทำให้เราคุมแถบสีไฮไลต์ได้เป๊ะ 100%

---

## 💡 6. วิธีต่อยอดและเพิ่มข้อมูล (How to Extend)

### ➕ วิธีเพิ่มเอกสารใหม่:
1. สร้างไฟล์ `MyNewDoc.md` ไปวางไว้ในโฟลเดอร์ `public/docs/`
2. สร้างเส้นทาง (Route) ใหม่ใน `App.jsx`:
   ```jsx
   <Route path="/docs/my-new-doc" element={<Layout><MarkdownViewer file="MyNewDoc.md" /></Layout>} />
   ```
3. เพิ่มเมนูใหม่ใน `Sidebar.jsx`:
   ```jsx
   <NavLink to="/docs/my-new-doc">📄 เอกสารใหม่</NavLink>
   ```

### ➕ วิธีเพิ่มแผนภาพใหม่ลงในระบบ:
1. สร้างคอมโพเนนต์ใหม่ เช่น `src/components/diagrams/MyFlow.jsx`
2. ไปที่ `src/components/MarkdownViewer.jsx` แล้วเพิ่ม Import:
   ```jsx
   import MyFlow from './diagrams/MyFlow';
   ```
3. เพิ่ม Switch Case ใน `DiagramRenderer`:
   ```jsx
   case 'my_flow': return <MyFlow />;
   ```
4. ในไฟล์ Markdown แค่พิมพ์ว่า:
   ````text
   ```diagram
   my_flow
   ```
   ````

---
*จัดทำขึ้นโดย AI Assistant 🤖 เพื่อใช้เป็นคู่มือการพัฒนาสำหรับโปรเจกต์ React + Bento Layout + Markdown Documentation*
