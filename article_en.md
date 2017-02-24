# Parsing docx with the help of XSLT

The task of handling office documents, namely docx documents, xlsx tables and pptx presentations is quite complicated. This article is about parsing, creating and editing documents using only XSLT and ZIP.
<cut />
Why?
docx is the most popular document format, so the ability to generate and parse this format  can always can be useful. The solution in a form of a ready-made library, can be inappropriate for several reasons:
- library may not exist
- you do not need another black box in your project 
- restrictions of the library: platforms, etc.
- licensing 
- processing speed

So, in this article I would use only basic tools for working with the docx documents.

## Docx structure
What is a docx document? A docx file is a zip archive which physically contains 2 types of files:
- xml files with `xml` and `rels` extensions
- media files (images, etc.)

And logically - 3 types of elements:
- Content Types - a type list of media files (e.g. png) used in the document and document parts (e.g. a document, a page header).
- Parts - separate document parts. For our document - it is document.xml, including xml documents and media files.
- Relationships identify document parts for links (e.g. communication between document section and page header), and also external parts are defined here (e.g. hyperlinks).


It is described in detail in the [ECMA-376: Office Open XML File Formats](http://www.ecma-international.org/publications/standards/Ecma-376.htm), the main part of it is a [PDF document](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fifth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip) consists of 5,000 pages and 2,000 more pages of bonus content.

## Minimal docx

[The simplest docx](https://github.com/eduard93/docx/releases/download/v1.0.0/minimal.docx) after unpacking looks like:

![image](https://habrastorage.org/files/ce5/f66/840/ce5f66840d3f4df484e083998829618c.PNG)

Let's take a [look](https://github.com/eduard93/docx/commit/5313b19d6b14392fee217f66afb11866fe738067) what it consists of.

#### [Content_Types].xml

It is located in document root and lists MIME types of document content:

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml"
              ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
```

#### _rels/.rels

The main list of document links. In this case, only one defined link - matching rId1 identifier and word/document.xml file - the main body of the document.
```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship 
        Id="rId1" 
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        Target="word/document.xml"/>
</Relationships>
```

#### word/document.xml
[Main document content](http://www.datypic.com/sc/ooxml/e-w_document.html).
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

Here:
- `<w:document>` - document itself
-  `<w:body>` - document body
- `<w:p>` - paragraph
- `<w:r>` - run (fragment) of the text
- `<w:t>` - text itself
- `<w:sectPr>` - page description

When you open this document in a text editor, you will see (document with) a single word `Test`.

#### word/_rels/document.xml.rels
It contains a list of links of `word/document.xml`. Name of link file is created from title of document part, to which it relates, and adding `rels` extension. A folder with link file called `_rels`, it is at the same level as a part to which it relates. There is no links in `word/document.xml`, so the file is empty:

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
```
Even if there is no links, file must exist.

## docx и Microsoft Word
[docx](https://github.com/eduard93/docx/releases/download/v1.0.0/word.docx) created with Microsoft Word or any other editor has [several additional files](https://github.com/eduard93/docx/commit/5313b19d6b14392fee217f66afb11866fe738067).

![image](https://habrastorage.org/files/585/503/504/58550350424d4977910f9424a4af3104.PNG)

Contents of files: 
- `docProps/core.xml` - the basic document metadata according to  [Open Packaging Conventions](https://en.wikipedia.org/wiki/Open_Packaging_Conventions) and Dublin Core  [[1]](http://dublincore.org/documents/dcmi-terms/), [[2]](http://dublincore.org/documents/dces/).
-  `docProps/app.xml` - [general information about document](http://www.datypic.com/sc/ooxml/e-extended-properties_Properties.html): number of pages, words, characters, application name in which document was created, etc.
- `word/settings.xml` - [settings for the current document](http://www.datypic.com/sc/ooxml/e-w_settings.html).
- `word/styles.xml` - [styles](http://www.datypic.com/sc/ooxml/e-w_styles.html) applied to the document.  Separate data from representation.
- `word/webSettings.xml` - HTML display [settings](http://www.datypic.com/sc/ooxml/e-w_webSettings.html) of document part and document conversion settings to HTML. 
- `word/fontTable.xml` - [list](http://www.datypic.com/sc/ooxml/e-w_fonts.html) of document fonts.
- `word/theme1.xml` - [theme](http://www.datypic.com/sc/ooxml/e-a_theme.html) (consists of color schemes, fonts, and formatting).

Complex documents can have much more parts.

## Reverse engineering docx

So, the initial task is to find out how any document fragment is stored in xml, then to create (or parse) such documents on their own. We need:
- Zip Archiver
- Library for XML formatting (Word gives XML without indents, one line)
- A tool for viewing diff between files, I use git and TortoiseGit

#### Tools
- For Windows:  [zip](http://gnuwin32.sourceforge.net/packages/zip.htm),  [unzip](http://gnuwin32.sourceforge.net/packages/unzip.htm), [libxml2](http://xmlsoft.org/downloads.html), [git](https://git-scm.com/download/win), [TortoiseGit](https://tortoisegit.org/download/)
- For Linux: ```apt-get install zip unzip libxml2 libxml2-utils git```

Also [scripts](https://github.com/eduard93/docx/commit/6b41b0e459329d62d0736aa6dc5a7b02e7398dcd) will be necessary for automatic archiving/dearching and XML formatting. 
Using on Windows:
-  `unpack file dir` - unpacks document `file` in folder `dir` and formats xml
-  `pack dir file` - pack folder `dir` in document `file`

Using on Linux is similar, but `./unpack.sh` instead of `unpack`, `pack` becomes `./pack`.

#### Use

Search changes:
1. Create a blank docx file in the editor. 
2. Unpack it using unpack in new folder. 
3. Commits new folder. 
4. Add to file from step 1. explored element (hyperlink, table, etc.). 
5. Unpack modified file into an existing folder. 
6. Explore diff, removing unnecessary changes (links permutation, order of namespaces, etc.). 
7. Packs folder and check opening of final file. 
8. Commit changed folder.

#### Example 1. Text selection bold

Finding of tag that defines text formatting in bold.

1.	Create `bold.docx` document with normal (not bold) text `Test`.
2.	Unpack it: `unpack bold.docx bold`.
3.	[Commit the result](https://github.com/eduard93/docx/commit/910ea3fb0f1667ce2722da491b27c4e12474c8ec).
4.	Select Test in bold.
5.	Unpack it: `unpack bold.docx bold`.
6.	Initially, the diff was as follows:

![diff](https://habrastorage.org/files/059/659/38c/05965938c8c64bbea20cb47fb5c6d457.PNG)
In detail: 

#### docProps/app.xml

```diff
@@ -1,9 +1,9 @@
-  <TotalTime>0</TotalTime>
+  <TotalTime>1</TotalTime>
```
Time change is not necessary.

#### docProps/core.xml
```diff
@@ -4,9 +4,9 @@
-  <cp:revision>1</cp:revision>
+  <cp:revision>2</cp:revision>
   <dcterms:created xsi:type="dcterms:W3CDTF">2017-02-07T19:37:00Z</dcterms:created>
-  <dcterms:modified xsi:type="dcterms:W3CDTF">2017-02-07T19:37:00Z</dcterms:modified>
+  <dcterms:modified xsi:type="dcterms:W3CDTF">2017-02-08T10:01:00Z</dcterms:modified>
```
Change document version and modification date is not necessary.

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

Changes in `w:rsidR` are unnecessary - it is inside information for Microsoft Word. A key change here:
```diff
         <w:rPr>
+          <w:b/>
```
in the paragraph with Test. Apparently element `<w:b/>` makes the text bold. Reserve this change and cancel the rest.

#### word/settings.xml

```diff
@@ -1,8 +1,9 @@
+  <w:proofState w:spelling="clean"/>
@@ -17,10 +18,11 @@
+    <w:rsid w:val="00F752CF"/>
```

It does not contain anything relating to the bold text. Cancel.

7 Pack a folder with 1m change (adding `<w:b/>`) and check that [document](https://github.com/eduard93/docx/releases/download/v1.0.0/bold.docx) opens and shows what was expected.
8 [Commit the change](https://github.com/eduard93/docx/commit/17f1dca258c44d87e8563b86a7e515b01bd4cee0).

#### Example 2. Footer

Complex example - adding footer.
[Initial commit](https://github.com/eduard93/docx/commit/0cd149e7cdab4e816a82a9128dbc5cfe89d74a97). Add footer text ‘123’ and unpack the document. Such initial diff looks like: 

![diff](https://habrastorage.org/files/478/e62/048/478e62048c12443481a00783f164bebe.PNG)

Immediately exclude changes in `docProps/app.xml` and `docProps/core.xml` – the same as in the first example.

#### [Content_Types].xml

```diff
@@ -4,10 +4,13 @@
   <Default Extension="xml" ContentType="application/xml"/>
   <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
+  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
+  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
+  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
```

footer looks clearly like what we need, but what we should do with footnotes and endnotes? Are they required by adding footer, or created them at the same time? The answer is not always easy, here are the basic ways: 
- View changes: are they connected with each other?
- Experiment
- Well, if you do not understand what`s happening: 

![Read the documentation](http://www.commitstrip.com/wp-content/uploads/2015/06/Strip-Lire-la-documentation-650-finalenglish.jpg)
Let`s go further.

#### word/_rels/document.xml.rels
Initial diff looks like:

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
We see that some of changes are due to fact that Word has changed link order, remove them:
```diff
@@ -3,6 +3,9 @@
+  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
+  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
+  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
```
footer, footnotes, endnotes appear again. All of them are connected with main document, take a look at it: 

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
There are only necessary changes – a clear link to footer from [sectPr](http://www.datypic.com/sc/ooxml/e-w_sectPr-3.html). There are no links to footnotes and endnotes in document, so we can assume links are not necessary.

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
Here are links to footnotes, endnotes which add them to document. 

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
We are interested in style changes, only if we are looking for how to change style. In this case, this change can be removed.

#### word/footer1.xml

Take a look at footer itself (some namespaces are omitted for readability, but in the document they should be):

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
Here is text: ‘123’. We need only one – remove the link to `<w:pStyle w:val="a6"/>`. 

The analysis of all the changes makes the following assumptions:
- footnotes and endnotes are unnecessary
- In `[Content_Types].xml` we need to add footer
- In `word/_rels/document.xml.rels` we need to add a link to footer
- In `word/document.xml` to `<w:sectPr>` tag we need to add `<w:footerReference>`

Reduce the diff to this set of changes:

![final diff](https://habrastorage.org/files/5d3/4fc/b84/5d34fcb8479244b198bc82507f61100a.PNG)

Then pack [document](https://github.com/eduard93/docx/releases/download/v1.0.0/footer.docx) and open it. If everything was done correctly, the document will be opened and there will be footer with text ‘123’. And here is the final [commit](https://github.com/eduard93/docx/commit/1f794a5cdba458b60466d8c1ca9a16e252b44e59).

Thus, the process of change detection is reduced to find a minimum set of changes sufficient to achieve the desired result.

## Practice

If we find necessary change, it is logical to proceed to the next stage, it could be any of:
- Create  docx
- Parse docx
- Convert docx

Here we need [XSLT](https://ru.wikipedia.org/wiki/XSLT) and [XPath](https://ru.wikipedia.org/wiki/XPath). 

Let's write a fairly simple conversion - replacement or addition of footer in the current document. I'm going to write in Caché ObjectScript, but even if you do not know this language - it does not matter. Basically, we will call XSLT and archiver, nothing more. So, let's start.

### Algorithm
Algorithm looks like:
1. Unpack the document
2. Add our footer 
3. Prescribe a link to it in `[Content_Types].xml` and `word/_rels/document.xml.rels` 
4. In `word/document.xml` to `<w:sectPr>` tag add `<w:footerReference>` tag or replace a link in it to our footer
5. Pack the document.

Let`s start.

#### Unpacking

In Caché ObjectScript it is possible to execute operating system commands using the function [$zf(-1, oscommand)](http://docs.intersystems.com/latest/csp/docbook/DocBook.UI.Page.cls?KEY=RCOS_fzf-1). Call unzip to unpack the document using [wrapper over $zf(-1)](https://github.com/intersystems-ru/Converter/blob/master/Converter/Common.cls.xml#L11):

```cos
/// Using %3 (unzip) unpack file %1 in folder %2
Parameter UNZIP = "%3 %1 -d %2";

/// Unpack archive source in folder targetDir
ClassMethod executeUnzip(source, targetDir) As %Status
{
    set timeout = 100
    set cmd = $$$FormatText(..#UNZIP, source, targetDir, ..getUnzip())
    return ..execute(cmd, timeout)
}

```

#### Creation of footer file

Input receives the footer text, we will write it to in.xml file:
```xml
<xml>TEST</xml>
```

In XSLT (file footer.xsl) we will create footer with text from xml tag (some namespaces are omitted, here is the [full list](https://github.com/intersystems-ru/Converter/blob/master/Converter/Footer.cls.xml#L327)): 
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

Call [XSLT converter](http://docs.intersystems.com/latest/csp/documatic/%25CSP.Documatic.cls?PAGE=CLASS&LIBRARY=%25SYS&CLASSNAME=%25XML.XSLT.Transformer#METHOD_TransformFile):
```cos
do ##class(%XML.XSLT.Transformer).TransformFile("in.xml", "footer.xsl", footer0.xml")    
```
The result is the footer file `footer0.xml`:
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

#### Add a footer link to a list of links of the main document

The link with `rId0` ID doesn't exist generally. However, you can use XPath to get the ID which does not exist. 
Add a link to `footer0.xml` with rId0 ID in `word/_rels/document.xml.rels`:

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

#### Specify links in document

Next, it is necessary in each `<w:sectPr>` tag add `<w:footerReference>` tag or replace a link in it to our footer. [It turns out](https://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.footerreference(v=office.14).aspx) that each of `<w:sectPr>` tag may have 3 `<w:footerReference>` tags - for the first page, even pages and the rest:

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

#### Add footer in `[Content_Types].xml`

Add in `[Content_Types].xml` information that `/word/footer0.xml` has a type of `application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml`:

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


#### As a result

Full code is  [published](https://github.com/intersystems-ru/Converter/blob/master/Converter/Footer.cls.xml). It works like this:
```cos
do ##class(Converter.Footer).modifyFooter("in.docx", "out.docx", "TEST")
```
Where:
- `in.docx` - original document
- `out.docx` - final document
- `TEST` - text which is added to footer

## Conclusions

Using only XSLT and ZIP, you can successfully work with docx documents, xlsx tables and pptx presentations.

## Open questions

1. Initially I wanted to use 7z instead of zip/unzip, as it is one tool and more common on Windows. However, I faced with such problem that documents packed 7z on Linux do not open in Microsoft Office. I tried to call a lot of [options](http://7zip.bugaco.com/7zip/MANUAL/switches/index.htm), but failed to achieve a positive result.
2. I`m looking for XSD with schemas ECMA-367 of version 5 and comments. The fifth XSD version is available for downloading on ECMA site. But it is difficult to understand it without any comments. The second XSD version with comments is available for downloading.

## Links
- [ECMA-376](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
- [docx discription](https://msdn.microsoft.com/en-us/library/aa338205.aspx)
- [Detailed article about docx](https://www.toptal.com/xml/an-informal-introduction-to-docx)
- [Repository with scripts](https://github.com/eduard93/docx)
- [Repository with footerconverter](https://github.com/intersystems-ru/Converter/)
