@echo off

rem это комментарий
rem echo — команда  предназначенная для отображения строки текста, @ знак нужен, чтобы сама команда не отображалась

rem Как добиться правильной работы пакетных файлов (*.bat, *.cmd), содержащих кириллические пути? 
rem Нужно сохранять текст в OEM кодировке (DOS-866). Это умеет делать текстовый редактор Wordpad

rem батник на копирование файлов 
rem меняет кодировку, чтобы видел русский текст
rem chcp 65001 и еще варинт, но не работает иногда chcp 1251 >nul
rem копирует (и заменяет если есть) файл в указанную папке (если ее нет - создает)
rem ПРИМЕР xcopy "P:\Сертификаты\Сертификаты для SAP\ARC Китай\Посуда из безцветного и цветного стекла (2).pdf"	 "C:\Users\roman.holodkov\Desktop\Новая папка\" /Y
rem имя компьютера
rem hostname 

rem tracert msk-sap-prd.bezant.local

rem открытие папки с файлами
rem start C:\Users
rem открытие папки "Документы" текущего пользователя
rem start %USERPROFILE%\Documents  
rem создание папки
rem mkdir %USERPROFILE%\Documents\ExcelMacro
rem перемещение move
rem выход из консоли
rem exit
rem >nul отключает вывод результата работы команды в консоль


chcp 65001 >nul
mkdir %USERPROFILE%\Documents\ExcelMacro >nul
echo ==============================================
rem ключи можно посмотреть xcopy /?   "%cd%" это текущая папка
xcopy "%cd%" %USERPROFILE%\Documents\ExcelMacro /Y /C /R /S >nul
echo == Открыта папка, в которую были скопированы макросы
start %USERPROFILE%\Documents\ExcelMacro
move "%USERPROFILE%\Documents\ExcelMacro\Остатки.xls - Ярлык.lnk" %USERPROFILE%\Desktop >nul
echo == Запущен экселевский файл, который поключает надстройку в Excel
echo (для установки макросов необходимо включить их в параметрах безопастности)
start %USERPROFILE%\Documents\ExcelMacro\AutoSetup.xlsm 
del "%USERPROFILE%\Documents\ExcelMacro\Автоустановка макросов.bat"
del %USERPROFILE%\Documents\ExcelMacro\AutoSetup.xlsm
echo == На рабочем столе был создан ярлык на файл Остатки (для его работы необходимо ODBC)
echo == Конец установки
echo ==============================================
pause




