@echo off
:: Set console encoding to UTF-8 to display Arabic correctly
chcp 65001 > nul
title وحدة التصاريح الأمنية - تشغيل النظام
color 0b

echo ==========================================================
echo               وحدة التصاريح الأمنية - نظام الاستمارات
echo ==========================================================
echo.
echo   [+] جاري تشغيل خادم الويب محلياً...
echo.
echo   [!] يمكنك الآن فتح المتصفح والدخول إلى النظام عبر:
echo       - رابط الموظف (عام): http://localhost:3000
echo       - لوحة تحكم المسؤول: http://localhost:3000/login.html
echo.
echo   [!] رمز دخول المسؤول الافتراضي: admin123
echo.
echo ==========================================================
echo.
node "%~dp0\server.js"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] حدث خطأ أثناء تشغيل السيرفر. يرجى التأكد من تثبيت الحزم بشكل صحيح.
    pause
)
