# FirstWord

**0.0.32**

Libre Office расширение для текстов с церковно-славянским языком.  

- Вставка внизу страницы первого слова со следующей страницы. Слово вставляется во врезку.
- Два первых слова, соединенных неразрывным пробелом, подчеркиванием и т.п., рассматриваются как одно.
- Начальные и конечные пробелы и знаки препинания (кроме конечной точки) удаляются.
- Учитываются особенности цся локали.
- Сохраняется некоторое форматирование (цвет, и в случае стиля для ч/б печати - жирность).
- Настроенные стили врезки и содержимого врезки (абзацный) предполагается в наличии, но при отсутствии создаются автоматически, и отчасти настраиваются. Далее их можно настроить под свои нужды.

Для текущей страницы:   

- Очистка содержимого врезки.
- Block/Unblock содержимого врезки.  
- Перемещение врезки вверх/вниз.
- Удаление врезки.
- Обновление содержимого врезки.

Запускается из LOffice для открытого документа.


**Для установки**  

Для Linux требуется установить компонент LibreOffice: *libreoffice-script-provider-python*  
``$ sudo apt-get install libreoffice-script-provider-python``  

**После установки**  

В OOffice функции доступны через _Addon's_ меню и панель.  


**Панель**  

На панели доступны кнопки вставки, удаления и обновления врезок с первым словом.     
![Вставка](src/Images/FW_16.png) &nbsp;&nbsp; ![Удаление](src/Images/FWRem_16.png) &nbsp;&nbsp; ![Обновление](src/Images/FWUpd_16.png)  
Обновить / Очистить / Удалить врезку только на текущей странице:  
![Обновить текущую врезку](src/Images/FWUpdCurr_16.png)&nbsp;&nbsp; ![Очистить текущую врезку](src/Images/FWClean_16.png) &nbsp;&nbsp; ![Удалить текущую врезку](src/Images/FWDel_16.png)  
Block/Unblock содержимого врезки.  
![Защитить содержимое врезки](src/Images/FWProtect_16.png) &nbsp;&nbsp; ![Разблокировать содержимое врезки](src/Images/FWUnProtect_16.png)  
Перемещение врезки вверх/вниз   
![Опустить содержимое врезки на 0.05](src/Images/FWDown_16.png) &nbsp;&nbsp; ![поднять содержимое врезки на 0.05](src/Images/FWUp_16.png)  

  


