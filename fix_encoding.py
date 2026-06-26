import os
path = r'C:\Users\nour\OneDrive\Desktop\RequestFlow\index.html'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()
print(f'Size: {os.path.getsize(path)}')
checks = ['اختيار القالب', 'نظام إدارة الطلبات', 'إدخال البيانات', 'DEFAULT_API_URL']
for c in checks:
    found = c in content
    print(f'  "{c}": {"OK" if found else "MISSING"}')

# Add explicit Content-Type
old = '<meta charset="UTF-8">'
new = '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">\n  <meta charset="UTF-8">'
if old in content:
    content = content.replace(old, new)
    print('Added Content-Type meta')
else:
    print('Warning: charset meta tag not found')

# Write back
with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Saved')
