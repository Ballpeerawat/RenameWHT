# 📄 **Withholding Tex Downloader - PDF Downloader & Rename**

โปรแกรมสำหรับ **ดาวน์โหลด** และ **เปลี่ยนชื่อไฟล์ PDF อัตโนมัติ** จากลิงก์ใน Google Sheet  
รองรับการทำงานกับเอกสารแบบพิมพ์จาก SCB หรือระบบอื่น ๆ ที่แนบลิงก์ภายใน Sheet

## 📥 ขั้นตอนการดาวน์โหลด

1. คลิกไฟล์ **Withholding Tax Downloader.exe** → ***Download***
2. ⚠ หากมีข้อความเตือน → คลิก ***Download anyway***
3. เมื่อดาวน์โหลดเสร็จ → เปิดโปรแกรม
4. จะมีหน้าต่าง **Windows protected your PC** ปรากฏขึ้น → ***คลิก More info***
5. จากนั้นคลิก **Run anyway** เพื่อเปิดโปรแกรม
6. **พร้อมใช้งาน!**

--- 
## 🔧 **ขั้นตอนการเตรียม Google Sheet (WHT_Print and Send)**

1. เปิดไฟล์ **Google Sheet: WHT_Print and Send**
2. ไปที่เมนู `Extensions > Apps Script`
3. วางโค้ดต่อไปนี้ในช่อง Script Editor:

    ```javascript
    function GetURL(input) {
      var myFormula = SpreadsheetApp.getActiveRange().getFormula();
      var myAddress = myFormula.replace('=GetURL(','').replace(')','');
      var myRange = SpreadsheetApp.getActiveSheet().getRange(myAddress);
      return myRange.getRichTextValue().getLinkUrl();
    }
    ```

4. กด **Save** แล้วคลิก **Run** ฟังก์ชัน (ครั้งแรกต้องกดยืนยันสิทธิ์)
5. กลับไปที่ Google Sheet แล้ว:
    - คลิกเซลล์ `Q8` ใส่ชื่อหัวคอลัมน์เป็น `URL`
    - ในเซลล์ `Q9` ใส่สูตร:

        ```excel
        =GetURL(J9)
        ```
        *(กรณีที่เซลล์ J9 มีข้อความ "WTI" ที่ฝังลิงก์ไว้)*

6. ลากสูตรลงให้ครบทุกแถวที่ต้องการ
7. ✅ เสร็จสิ้นการเตรียม Sheet

---

## 📥 **การใช้งานโปรแกรม Withholding Tex Downloader**

1. เปิดโปรแกรม **Withholding Tex Downloader**
2. เปิด Google Sheet และ **คัดลอกลิงก์ -> Link for program**
3. วางลิงก์ในช่อง **Sheet URL** ของโปรแกรม
4. คลิก **Start Download**
5. โปรแกรมจะทำงานโดยอัตโนมัติ:
    - เปิดลิงก์เอกสารทีละรายการ
    - คลิกปุ่ม **“พิมพ์เอกสาร”**
    - ดาวน์โหลด PDF และ **เปลี่ยนชื่อไฟล์ตามข้อมูลใน Sheet**

6. ไฟล์ PDF จะถูกจัดเก็บไว้ในโฟลเดอร์:  
   `Downloads/Download WHT`

7. ✅ เสร็จสิ้นการดาวน์โหลด

---

## ⚠️ **หมายเหตุสำคัญ**

- 🔓 **Google Sheet ต้องเปิดสิทธิ์การแชร์แบบ "Anyone with the link: Editor"**  
  *(สามารถปิดการแชร์ได้หลังจากดาวน์โหลดเสร็จ)*
