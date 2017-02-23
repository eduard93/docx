# Как я разбирал docx с помощью XSLT

Задача обработки документов в формате docx, а также таблиц xlsx и презентаций pptx является весьма нетривиальной. В этой статье расскажу как научиться парсить, создавать и обрабатывать такие документы используя только XSLT и ZIP архиватор.
<cut />
## Зачем?
docx - самый популярный формат документов, поэтому задача отдавать информацию пользователю в этом формате всегда может возникнуть. Один из вариантов решения этой проблемы - использование готовой библиотеки, может не подходить по ряду причин:
- библиотеки может просто не существовать
- в проекте не нужен ещё один чёрный ящик
- ограничения библиотеки по платформам и т.п.
- проблемы с лицензированием
- скорость работы

Поэтому в этой статье будем использовать только самые базовые инструменты для работы с docx документом.

## Структура docx
Для начала разоберёмся с тем, что собой представляет docx документ. docx это zip архив который физически содержит 2 типа файлов:
- xml файлы с расширениями `xml` и `rels`
- медиа файлы  (изображения и т.п.)

А логически - 3 вида элементов:
- Типы (Content Types) - список типов медиа файлов (например png) встречающихся в документе и типов частей документов (например документ, верхний колонтитул).
- Части (Parts) - отдельные части документа, для нашего документа это document.xml, сюда входят как xml документы так и медиа файлы.
- Связи (Relationships) идентифицируют части документа для ссылок (например связь между разделом документа и колонтитулом), а также тут определены внешние части (например гиперссылки).

Они подробно описаны в стандарте [ECMA-376: Office Open XML File Formats](http://www.ecma-international.org/publications/standards/Ecma-376.htm), основная часть которого - [PDF документ](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fifth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip) на 5000 страниц, и ещё 2000 страниц бонусного контента.

## Минимальный docx

[Простейший docx](https://github.com/eduard93/docx/releases/download/v1.0.0/minimal.docx) после распаковки выглядит следующим образом

![image](https://habrastorage.org/files/ce5/f66/840/ce5f66840d3f4df484e083998829618c.PNG)

Давайте [посмотрим](https://github.com/eduard93/docx/commit/5313b19d6b14392fee217f66afb11866fe738067) из чего он состоит.

#### [Content_Types].xml

Находится в корне документа и перечисляет MIME типы содержимого документа:

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
```

#### _rels/.rels

Главный список связей документа. В данном случае определена всего одна связь - сопоставление с идентификатором rId1 и файлом word/document.xml - основным телом документа.
```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship 
        Id="rId1" 
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        Target="word/document.xml"/>
</Relationships>
```

#### word/document.xml
[Основное содержимое документа](http://www.datypic.com/sc/ooxml/e-w_document.html).
<spoiler title="word/document.xml">
```xml
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14 wp14">
    <w:body>
        <w:p w:rsidR="005F670F" w:rsidRDefault="005F79F5">
            <w:r>
                <w:t>Test</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
        </w:p>
        <w:sectPr w:rsidR="005F670F">
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" 
                     w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
            <w:docGrid w:linePitch="360"/>
        </w:sectPr>
    </w:body>
</w:document>
```
</spoiler>

Здесь:
- `<w:document>` - сам документ
-  `<w:body>` - тело документа
- `<w:p>` - параграф
- `<w:r>` - run (фрагмент) текста
- `<w:t>` - сам текст
- `<w:sectPr>` - описание страницы

Если открыть этот документ в текстовом редакторе, то увидим документ из одного слова `Test`.

#### word/_rels/document.xml.rels
Здесь содержится список связей части `word/document.xml`. Название файла связей создаётся из названия части документа к которой он относится и добавления к нему расширения `rels`. Папка с файлом связей называется `_rels` и находится на том же уровне, что и часть к которой он относится. Так как связей в `word/document.xml` никаких нет то и в файле пусто:

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
```
Даже если связей нет, этот файл должен существовать.

## docx и Microsoft Word
[docx](https://github.com/eduard93/docx/releases/download/v1.0.0/word.docx) созданный с помощью Microsoft Word, да в принципе и с помощью любого другого редактора имеет [несколько дополнительных файлов](https://github.com/eduard93/docx/commit/5313b19d6b14392fee217f66afb11866fe738067).

![image](https://habrastorage.org/files/585/503/504/58550350424d4977910f9424a4af3104.PNG)

Вот что в них содержится:
- `docProps/core.xml` - основные метаданные документа согласно [Open Packaging Conventions](https://en.wikipedia.org/wiki/Open_Packaging_Conventions) и Dublin Core [[1]](http://dublincore.org/documents/dcmi-terms/), [[2]](http://dublincore.org/documents/dces/).
-  `docProps/app.xml` - [общая информация о документе](http://www.datypic.com/sc/ooxml/e-extended-properties_Properties.html): количество страниц, слов, символов, название приложения в котором был создан документ и т.п.
- `word/settings.xml` - [настройки относящиеся к текущему документу](http://www.datypic.com/sc/ooxml/e-w_settings.html).
- `word/styles.xml` - [стили](http://www.datypic.com/sc/ooxml/e-w_styles.html) применимые к документу. Отделяют данные от представления.
- `word/webSettings.xml` - [настройки](http://www.datypic.com/sc/ooxml/e-w_webSettings.html) отображения HTML частей документа и настройки того, как конвертировать документ в HTML.
- `word/fontTable.xml` - [список](http://www.datypic.com/sc/ooxml/e-w_fonts.html) шрифтов используемых в документе.
- `word/theme1.xml` - [тема](http://www.datypic.com/sc/ooxml/e-a_theme.html) (состоит из цветовой схемы, шрифтов и форматирования).

В сложных документах частей может быть гораздо больше.

## Реверс-инжиниринг docx

Итак, первоначальная задача - узнать как какой-либо фрагмент документа хранится в xml, чтобы потом создавать (или парсить) подобные документы самостоятельно. Для этого нам понадобятся:
- Архиватор zip
- Библиотека для форматирования XML (Word выдаёт XML без отступов, одной строкой)
- Средство для просмотра diff между файлами, я буду использовать git и TortoiseGit

#### Инструменты
- Под Windows: [zip](http://gnuwin32.sourceforge.net/packages/zip.htm),  [unzip](http://gnuwin32.sourceforge.net/packages/unzip.htm), [libxml2](http://xmlsoft.org/downloads.html), [git](https://git-scm.com/download/win), [TortoiseGit](https://tortoisegit.org/download/)
- Под Linux: ```apt-get install zip unzip libxml2 libxml2-utils git```

Также понадобятся [скрипты](https://github.com/eduard93/docx/commit/6b41b0e459329d62d0736aa6dc5a7b02e7398dcd) для автоматического (раз)архивирования и форматирования XML.
Использование под Windows:
-  `unpack file dir` - распаковывает документ `file` в папку `dir` и форматирует xml
-  `pack dir file` - запаковывает папку `dir` в документ `file` 

Использование под Linux аналогично, только `./unpack.sh` вместо `unpack`, а `pack` становится `./pack`.

#### Использование

Поиск изменений происходит следующим образом:
1. Создаём пустой docx файл в редакторе.
2. Распаковываем его с помощью `unpack` в новую папку.
3. Коммитим новую папку.
4. Добавляем в файл из п. 1. изучаемый элемент (гиперссылку, таблицу и т.д.).
5. Распаковываем изменённый файл в уже существующую папку.
6. Изучаем diff, убирая ненужные изменения (перестановки связей, порядок пространств имён и т.п.).
7. Запаковываем папку и проверяем что получившийся файл открывается.
8. Коммитим изменённую папку.

#### Пример 1. Выделение текста жирным

Посмотрим на практике, как найти тег который определяет форматирование текста жирным шрифтом.

1. Создаём документ `bold.docx` с обычным (не жирным) текстом Test.
2. Распаковываем его: `unpack bold.docx bold`.
3. [Коммитим результат](https://github.com/eduard93/docx/commit/910ea3fb0f1667ce2722da491b27c4e12474c8ec).
4. Выделяем текст Test жирным.
5. Распаковываем `unpack bold.docx bold`.
6. Изначально diff выглядел следующим образом: 

![diff](https://habrastorage.org/files/059/659/38c/05965938c8c64bbea20cb47fb5c6d457.PNG)
Рассмотрим его подробно:

#### docProps/app.xml

```diff
@@ -1,9 +1,9 @@
-  <TotalTime>0</TotalTime>
+  <TotalTime>1</TotalTime>
```
Изменение времени нам не нужно.

#### docProps/core.xml
```diff
@@ -4,9 +4,9 @@
-  <cp:revision>1</cp:revision>
+  <cp:revision>2</cp:revision>
   <dcterms:created xsi:type="dcterms:W3CDTF">2017-02-07T19:37:00Z</dcterms:created>
-  <dcterms:modified xsi:type="dcterms:W3CDTF">2017-02-07T19:37:00Z</dcterms:modified>
+  <dcterms:modified xsi:type="dcterms:W3CDTF">2017-02-08T10:01:00Z</dcterms:modified>
```
Изменение версии документа и даты модификации нас также не интересует.
 
#### word/document.xml
<spoiler title="diff">
```diff
@@ -1,24 +1,26 @@
    <w:body>
-    <w:p w:rsidR="0076695C" w:rsidRPr="00290C70" w:rsidRDefault="00290C70">
+    <w:p w:rsidR="0076695C" w:rsidRPr="00F752CF" w:rsidRDefault="00290C70">
       <w:pPr>
         <w:rPr>
+          <w:b/>
           <w:lang w:val="en-US"/>
         </w:rPr>
       </w:pPr>
-      <w:r>
+      <w:r w:rsidRPr="00F752CF">
         <w:rPr>
+          <w:b/>
           <w:lang w:val="en-US"/>
         </w:rPr>
         <w:t>Test</w:t>
       </w:r>
       <w:bookmarkStart w:id="0" w:name="_GoBack"/>
       <w:bookmarkEnd w:id="0"/>
     </w:p>
-    <w:sectPr w:rsidR="0076695C" w:rsidRPr="00290C70">
+    <w:sectPr w:rsidR="0076695C" w:rsidRPr="00F752CF">
```
</spoiler>

Изменения в  `w:rsidR` не интересны - это внутренняя информация для Microsoft Word. Ключевое изменение тут
```diff
         <w:rPr>
+          <w:b/>
```
в параграфе с Test.  Видимо элемент `<w:b/>` и делает текст жирным. Оставляем это изменение и отменяем остальные.

#### word/settings.xml

```diff
@@ -1,8 +1,9 @@
+  <w:proofState w:spelling="clean"/>
@@ -17,10 +18,11 @@
+    <w:rsid w:val="00F752CF"/>
```

Также не содержит ничего относящегося к жирному тексту. Отменяем.

7 Запаковываем папку с 1м изменением (добавлением `<w:b/>`) и проверяем что [документ](https://github.com/eduard93/docx/releases/download/v1.0.0/bold.docx) открывается и показывает то, что ожидалось.
8 [Коммитим изменение](https://github.com/eduard93/docx/commit/17f1dca258c44d87e8563b86a7e515b01bd4cee0).

#### Пример 2. Нижний колонтитул

Теперь разберём пример посложнее - добавление нижнего колонтитула.
[Вот первоначальный коммит](https://github.com/eduard93/docx/commit/0cd149e7cdab4e816a82a9128dbc5cfe89d74a97). Добавляем нижний колонтитул с текстом 123 и распаковываем документ. Такой diff получается первоначально:

![diff](https://habrastorage.org/files/478/e62/048/478e62048c12443481a00783f164bebe.PNG)

Сразу же исключаем изменения в `docProps/app.xml` и `docProps/core.xml` - там тоже самое, что и в первом примере.

#### [Content_Types].xml

```diff
@@ -4,10 +4,13 @@
   <Default Extension="xml" ContentType="application/xml"/>
   <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
+  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
+  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
+  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
```

footer явно выглядит как то, что нам нужно, но что делать с footnotes и endnotes? Являются ли они обязательными при добавлении нижнего колонтитула или их создали заодно? Ответить на этот вопрос не всегда просто, вот основные пути:
- Посмотреть, связаны ли изменения друг с другом
- Экспериментировать
- Ну а если совсем не понятно что происходит:

![Читать документацию](http://www.commitstrip.com/wp-content/uploads/2015/06/Strip-Lire-la-documentation-650-finalenglish.jpg)
Идём пока что дальше.

#### word/_rels/document.xml.rels
Изначально diff выглядит вот так:

<spoiler title="diff">
```diff
@@ -1,8 +1,11 @@
 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
+  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
   <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
+  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
   <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
   <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
-  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
-  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
+  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
+  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
+  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
 </Relationships>
```
</spoiler>
Видно, что часть изменений связана с тем, что Word изменил порядок связей, уберём их:
```diff
@@ -3,6 +3,9 @@
+  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
+  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
+  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
```
Опять появляются footer, footnotes, endnotes. Все они связаны с основным документом, перейдём к нему:

#### word/document.xml
```diff
@@ -15,10 +15,11 @@
       </w:r>
       <w:bookmarkStart w:id="0" w:name="_GoBack"/>
       <w:bookmarkEnd w:id="0"/>
     </w:p>
     <w:sectPr w:rsidR="0076695C" w:rsidRPr="00290C70">
+      <w:footerReference w:type="default" r:id="rId6"/>
       <w:pgSz w:w="11906" w:h="16838"/>
       <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
       <w:cols w:space="708"/>
       <w:docGrid w:linePitch="360"/>
     </w:sectPr>
```
Редкий случай когда есть только нужные изменения. Видна явная ссылка на footer из [sectPr](http://www.datypic.com/sc/ooxml/e-w_sectPr-3.html). А так как ссылок в документе на footnotes и endnotes нет, то можно предположить что они нам не понадобятся.

#### word/settings.xml
<spoiler title="diff">

```diff
@@ -1,19 +1,30 @@
 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15">
   <w:zoom w:percent="100"/>
+  <w:proofState w:spelling="clean"/>
   <w:defaultTabStop w:val="708"/>
   <w:characterSpacingControl w:val="doNotCompress"/>
+  <w:footnotePr>
+    <w:footnote w:id="-1"/>
+    <w:footnote w:id="0"/>
+  </w:footnotePr>
+  <w:endnotePr>
+    <w:endnote w:id="-1"/>
+    <w:endnote w:id="0"/>
+  </w:endnotePr>
   <w:compat>
     <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
     <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
     <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
     <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
     <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
   </w:compat>
   <w:rsids>
     <w:rsidRoot w:val="00290C70"/>
+    <w:rsid w:val="000A7B7B"/>
+    <w:rsid w:val="001B0DE6"/>
```
</spoiler>
А вот и появились ссылки на footnotes, endnotes добавляющие их в документ.

#### word/styles.xml

<spoiler title="diff">
```diff
@@ -480,6 +480,50 @@
       <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
       <w:b/>
       <w:sz w:val="28"/>
     </w:rPr>
   </w:style>
+  <w:style w:type="paragraph" w:styleId="a4">
+    <w:name w:val="header"/>
+    <w:basedOn w:val="a"/>
+    <w:link w:val="a5"/>
+    <w:uiPriority w:val="99"/>
+    <w:unhideWhenUsed/>
+    <w:rsid w:val="000A7B7B"/>
+    <w:pPr>
+      <w:tabs>
+        <w:tab w:val="center" w:pos="4677"/>
+        <w:tab w:val="right" w:pos="9355"/>
+      </w:tabs>
+      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
+    </w:pPr>
+  </w:style>
+  <w:style w:type="character" w:customStyle="1" w:styleId="a5">
+    <w:name w:val="Верхний колонтитул Знак"/>
+    <w:basedOn w:val="a0"/>
+    <w:link w:val="a4"/>
+    <w:uiPriority w:val="99"/>
+    <w:rsid w:val="000A7B7B"/>
+  </w:style>
+  <w:style w:type="paragraph" w:styleId="a6">
+    <w:name w:val="footer"/>
+    <w:basedOn w:val="a"/>
+    <w:link w:val="a7"/>
+    <w:uiPriority w:val="99"/>
+    <w:unhideWhenUsed/>
+    <w:rsid w:val="000A7B7B"/>
+    <w:pPr>
+      <w:tabs>
+        <w:tab w:val="center" w:pos="4677"/>
+        <w:tab w:val="right" w:pos="9355"/>
+      </w:tabs>
+      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
+    </w:pPr>
+  </w:style>
+  <w:style w:type="character" w:customStyle="1" w:styleId="a7">
+    <w:name w:val="Нижний колонтитул Знак"/>
+    <w:basedOn w:val="a0"/>
+    <w:link w:val="a6"/>
+    <w:uiPriority w:val="99"/>
+    <w:rsid w:val="000A7B7B"/>
+  </w:style>
 </w:styles>
```
</spoiler>
Изменения в стилях нас интересуют только если мы ищем как поменять стиль. В данном случае это изменение можно убрать.

#### word/footer1.xml

Посмотрим теперь собственно на сам нижний колонтитул (часть пространств имён опущена для читабельности, но в документе они должны быть): 

```xml
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p w:rsidR="000A7B7B" w:rsidRDefault="000A7B7B">
    <w:pPr>
      <w:pStyle w:val="a6"/>
    </w:pPr>
    <w:r>
      <w:t>123</w:t>
    </w:r>
  </w:p>
</w:ftr>
```
Тут виден текст 123. Единственное, что надо исправить - убрать ссылку на `<w:pStyle w:val="a6"/>`.

В результате анализа всех изменений делаем следующие предположения:
- footnotes и endnotes не нужны
- В `[Content_Types].xml` надо добавить footer
- В `word/_rels/document.xml.rels` надо добавить ссылку на footer
- В `word/document.xml` в тег `<w:sectPr>` надо добавить `<w:footerReference>`

Уменьшаем diff до этого набора изменений:

![final diff](https://habrastorage.org/files/5d3/4fc/b84/5d34fcb8479244b198bc82507f61100a.PNG)

Затем запаковываем [документ](https://github.com/eduard93/docx/releases/download/v1.0.0/footer.docx) и открываем его. 
Если всё сделано правильно, то документ откроется и в нём будет нижний колонтитул с текстом 123. А вот и итоговый [коммит](https://github.com/eduard93/docx/commit/1f794a5cdba458b60466d8c1ca9a16e252b44e59).

Таким образом процесс поиска изменений сводится к поиску минимального набора изменений, достаточного для достижения заданного результата.

## Практика

Найдя интересующее нас изменение, логично перейти к следующему этапу, это может быть что-либо из:
- Создания docx
- Парсинг docx
- Преобразования docx  

Тут нам потребуются знания [XSLT](https://ru.wikipedia.org/wiki/XSLT) и [XPath](https://ru.wikipedia.org/wiki/XPath). 

Давайте напишем достаточно простое преобразование - замену или добавление нижнего колонтитула в существующий документ. Писать я буду на языке Caché ObjectScript, но даже если вы не знаете - не беда. В основном будем вызовать XSLT и архиватор. Ничего более. Итак, приступим.

### Алгоритм
Алгоритм выглядит следующим образом:
1. Распаковываем документ
2. Добавляем наш нижний колонтитул
3. Прописываем ссылку на него в `[Content_Types].xml` и `word/_rels/document.xml.rels`
4. В `word/document.xml` в тег `<w:sectPr>` добавляем тег `<w:footerReference>` или заменяем в нём ссылку на наш нижний колонтитул.
5. Запаковываем документ

Приступим.

#### Распаковка

В Caché ObjectScript есть возможность выполнять команды ОС с помощью функции [$zf(-1, oscommand)](http://docs.intersystems.com/latest/csp/docbook/DocBook.UI.Page.cls?KEY=RCOS_fzf-1). Вызовем unzip для распаковки документа с помощью [обёртки над $zf(-1)](https://github.com/intersystems-ru/Converter/blob/master/Converter/Common.cls.xml#L11):

```cos
/// Используя %3 (unzip) распаковать файл %1 в папку %2
Parameter UNZIP = "%3 %1 -d %2";

/// Распаковать архив source в папку targetDir
ClassMethod executeUnzip(source, targetDir) As %Status
{
    set timeout = 100
    set cmd = $$$FormatText(..#UNZIP, source, targetDir, ..getUnzip())
    return ..execute(cmd, timeout)
}

```

#### Создаём файл нижнего колонтитула

На вход поступает текст нижнего колонтитула, запишем его в файл in.xml:
```xml
<xml>TEST</xml>
```

В XSLT (файл - footer.xsl) будем создавать нижний колонтитул с текстом из тега xml (часть пространств имён опущена, вот [полный список](https://github.com/intersystems-ru/Converter/blob/master/Converter/Footer.cls.xml#L327)): 
```xml
<xsl:stylesheet 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships" version="1.0">
    <xsl:output method="xml" omit-xml-declaration="no" indent="yes" standalone="yes"/>
    <xsl:template match="/">

        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p>
                <w:r>
                    <w:rPr>
                        <w:lang w:val="en-US"/>
                    </w:rPr>
                    <w:t>
                        <xsl:value-of select="//xml/text()"/>
                    </w:t>
                </w:r>
            </w:p>
        </w:ftr>
    </xsl:template>
</xsl:stylesheet>
```

Теперь вызовем [XSLT преобразователь](http://docs.intersystems.com/latest/csp/documatic/%25CSP.Documatic.cls?PAGE=CLASS&LIBRARY=%25SYS&CLASSNAME=%25XML.XSLT.Transformer#METHOD_TransformFile):
```cos
do ##class(%XML.XSLT.Transformer).TransformFile("in.xml", "footer.xsl", footer0.xml")    
```
В результате получится файл нижнего колонтитула `footer0.xml`: 
```xml
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:r>
            <w:rPr>
                <w:lang w:val="en-US"/>
            </w:rPr>
            <w:t>TEST</w:t>
        </w:r>
    </w:p>
</w:ftr>
```

#### Добавляем ссылку на колонтитул в список связей основного документа

Сссылки с идентификатором `rId0` как правило не существует. Впрочем можно использовать XPath для получения идентификатора которого точно не существует. 
Добавляем ссылку на `footer0.xml` c идентификатором rId0   в `word/_rels/document.xml.rels`:

<spoiler title="XSLT">
```xml
<xsl:stylesheet  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"  version="1.0">
    <xsl:output method="xml" omit-xml-declaration="yes" indent="no"  />
    <xsl:param name="new">
        <Relationship 
           Id="rId0" 
           Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" 
           Target="footer0.xml"/>
    </xsl:param>

    <xsl:template match="/*">
        <xsl:copy>
            <xsl:copy-of select="$new"/>
            <xsl:copy-of select="@* | node()"/>
        </xsl:copy>
    </xsl:template>
</xsl:stylesheet>
```
</spoiler>

#### Прописываем ссылки в документе

Далее надо в каждый тег `<w:sectPr>` добавить тег `<w:footerReference>` или заменить в нём ссылку на наш нижний колонтитул. [Оказалось](https://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.footerreference(v=office.14).aspx), что у каждого тега `<w:sectPr>` может быть 3 тега `<w:footerReference>` - для первой страницы, четных страниц и всего остального:
<spoiler title="XSLT">
```xml
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
version="1.0">
    <xsl:output method="xml" omit-xml-declaration="yes" indent="yes" />
    <xsl:template match="//@* | //node()">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <xsl:apply-templates select="node()"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="//w:sectPr">
        <xsl:element name="{name()}" namespace="{namespace-uri()}">
            <xsl:copy-of select="./namespace::*"/>
            <xsl:apply-templates select="@*"/>
            <xsl:copy-of select="./*[local-name() != 'footerReference']"/>
            <w:footerReference w:type="default" r:id="rId0"/>
            <w:footerReference w:type="first" r:id="rId0"/>
            <w:footerReference w:type="even" r:id="rId0"/>
        </xsl:element>
    </xsl:template>
</xsl:stylesheet>
```
</spoiler>

#### Добавляем колонтитул в `[Content_Types].xml`

Добавляем в `[Content_Types].xml` информацию о том, что `/word/footer0.xml` имеет тип `application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml`:

<spoiler title="XSLT">
```xml
<xsl:stylesheet  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  xmlns="http://schemas.openxmlformats.org/package/2006/content-types"  version="1.0">
    <xsl:output method="xml" omit-xml-declaration="yes" indent="no"  />
    <xsl:param name="new">
        <Override 
         PartName="/word/footer0.xml" 
         ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
    </xsl:param>

    <xsl:template match="/*">
        <xsl:copy>
            <xsl:copy-of select="@* | node()"/> 
            <xsl:copy-of select="$new"/>
        </xsl:copy>
    </xsl:template>
</xsl:stylesheet>
```
</spoiler>


#### В результате

Весь код [опубликован](https://github.com/intersystems-ru/Converter/blob/master/Converter/Footer.cls.xml). Работает он так:
```cos
do ##class(Converter.Footer).modifyFooter("in.docx", "out.docx", "TEST")
```
Где:
- `in.docx` - исходный документ
- `out.docx` - выходящий документ
- `TEST` - текст, который добавляется в нижний колонтитул

## Выводы

Используя только XSLT и ZIP можно успешно работать с документами docx, таблицами xlsx и презентациями pptx.

## Открытые вопросы

1. Изначально хотел использовать 7z вместо zip/unzip т..к. это одна утилита и она более распространена на Windows. Однако я столкнулся с такой проблемой, что документы запакованные 7z под Linux не открываются в Microsoft Office. Я попробовал достаточно много [вариантов](http://7zip.bugaco.com/7zip/MANUAL/switches/index.htm) вызова, однако положительного результата добиться не удалось.
2. Ищу XSD со схемами ECMA-376 версии 5 и комментариями. XSD версии 5 без комментариев доступен к загрузке на сайте ECMA, но без комментариев в нём сложно разобраться. XSD версии 2 с комментариями доступен к загрузке.

## Ссылки
- [ECMA-376](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
- [Описание docx](https://msdn.microsoft.com/en-us/library/aa338205.aspx)
- [Подробная статья про docx](https://www.toptal.com/xml/an-informal-introduction-to-docx)
- [Репозиторий со скриптами](https://github.com/eduard93/docx)
- [Репозиторий с преобразователем нижнего колонтитула](https://github.com/intersystems-ru/Converter/)