# RequestFlow

نظام إدارة الطلبات - جمعية صناع الحياة المنوفية  
إرسال الطلبات عبر فورم HTML → إنشاء مستند في Google Drive + تسجيل في Google Sheets

## الملفات

| الملف | الوظيفة |
|-------|---------|
| `index.html` | واجهة المستخدم (GitHub Pages) - الفورم المباشر |
| `Code.gs` | Google Apps Script backend (API) |

## رابط الفورم
**[https://kareemshair2.github.io/RequestFlow](https://kareemshair2.github.io/RequestFlow)**

## طريقة النشر (مرة واحدة)

### 1. ربط Google Apps Script
1. افتح الشيت: https://docs.google.com/spreadsheets/d/1x7UhPnFWM8RHYacFoa9tHG5yoWA7cRDHZg5xFkv-tW4
2. **Extensions > Apps Script** > سمّي المشروع "RequestFlow"
3. انسخ محتوى `Code.gs` كاملاً في ملف `Code.gs`
4. **Resources > Cloud Platform Project** > فعّل **Drive API**
5. شغّل دالة `setupSheet()` من المحرر (حتى تنشئ التبوت)
6. **Deploy > New Deployment > Web App**:
   - Execute as: **Me**
   - Access: **Anyone with link**
   - انسخ الرابط الناتج (هو دا API URL)

### 2. تجهيز قوالب Drive
- حمل ملفات القوالب من مجلد `templates/modified/` إلى مجلد `تمبلت الطلبات` في Drive

### 3. الفورم جاهز!
- الفورم على GitHub Pages مربوط تلقائياً بالـ API
- ابدأ من: https://kareemshair2.github.io/RequestFlow

## النظام بيشتغل إزاي؟
1. المستخدم بيفتح الفورم
2. بيختار القالب (طلب شراء / إيصال استلام / كشف توزيع)
3. بيملأ البيانات وبيضغط **تقديم الطلب**
4. النظام:
   - ينشئ مستند في Drive (جوه مجلد البند > تاريخ - عنوان)
   - يسجل البيانات في Sheet (تاب للبند + تاب رئيسي `all`)

## License
MIT
