# OpenXmlWordCreater

## 說明

將設定書籤的 Word (*.docx) 範本透過 .NET Core OpenXml 套件將書籤帶入對應的文字內容

> 套件 DocumentFormat.OpenXml

本範例程式碼使用 .NET Standard 2.1 執行

## 如何操作

變更範本 (*.docx) 檔案與書籤位置，並給予合適的命名

建立新的範本資料物件 (class) 並修改範本物件與書籤對應方法

> ... ToKeyValues(object item) ...

## 在 Word 檔案設定書籤

![](https://raw.githubusercontent.com/txstudio/OpenXmlWordCreater/master/images/add-bookmark-location.gif)

![](https://raw.githubusercontent.com/txstudio/OpenXmlWordCreater/master/images/registration-from-in-bookmark.gif)

## 注意事項

在不調整原始程式碼的情況下，會有下列限制

請使用 "_" 字元標記要替換的書籤 (bookmark)

> 例： "UserName_"

書籤需設定在「表格」的儲存格中

> 不然文字的定位可能會錯誤

預設字型會設定如下

> 中文：標楷體
> 非中文：Times New Roman

## 後紀

若直接使用 XML Replace 嘗試取代 OpenXml 的數值的話，可能會因為標籤過於複雜無法取代成功

## 參考資料

https://stackoverflow.com/questions/21104210/reading-word-bookmarks-using-open-xml-and-c-sharp

https://docs.microsoft.com/zh-tw/office/open-xml/open-xml-sdk
