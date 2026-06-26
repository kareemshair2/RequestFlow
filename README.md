# RequestFlow

نظام إدارة الطلبات - جمعية صناع الحياة المنوفية  
إرسال الطلبات عبر فورم HTML → إنشاء مستند Word في Drive + تسجيل في Google Sheets

## الملفات

| الملف | الوظيفة |
|-------|---------|
| `Code.gs` | Google Apps Script backend - يتولى إنشاء الملفات والمجلدات والتسجيل في الشيت |
| `Index.html` | واجهة المستخدم الرئيسية (الفورم) |
| `Stylesheet.html` | تنسيقات CSS |
| `JavaScript.html` | منطق الفورم (إضافة صفوف، التحقق من البيانات) |
| `add_placeholders.py` | سكربت Python لإضافة `{{placeholders}}` لملفات القوالب |

## طريقة النشر

### 1. تجهيز القوالب
- شغّل `add_placeholders.py` عشان تضيف placeholders لملفات القوالب
- ارفع الملفات الناتجة من مجلد `modified/` إلى Google Drive

### 2. ربط Google Sheets
1. افتح الشيت: `Extensions > Apps Script`
2. أنشئ مشروع جديد وسمّه "RequestFlow"
3. انسخ محتوى الملفات الأربعة (Code.gs, Index.html, Stylesheet.html, JavaScript.html) في الملفات المقابلة
4. فعّل Drive API من Resources > Cloud Platform Project
5. شغّل `setupSheet()` من المحرر عشان تنشئ التبوت
6. انشر: Deploy > New Deployment > Web App (Execute as: Me, Access: Anyone)

### 3. استخدام الفورم
- خذ رابط Web App ووزعه على المستخدمين
- المستخدم يختار القالب، يملأ البيانات، ويضغط "تقديم الطلب"
- النظام ينشئ المستند في Drive ويسجل البيانات في Sheet

## القوالب المتوفرة
- **طلب شراء (سلعة/خدمة)** - مع جدول أصناف
- **إيصال استلام منحة عينية** - إقرار استلام
- **كشف توزيع** - كشف بأسماء المستفيدين

## License
MIT
