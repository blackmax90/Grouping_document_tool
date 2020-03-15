# Grouping_document_tool

## INTRODUCTION
This tools can group the documents(MS Office - .docx, .doc, .pptx, .ppt, .xlsx, .xls) by selecting the desired feature (Company, author, last saved by, layout, RSID (.docx only), property information, and everything). Considering the efficiency of the investigation, the tool can easily view only the grouped document files by copying the grouped document files and storing them in folders. In addition, the details of the grouped document files (filename, company, author, last modified by, etc.) are saved in the .csv file. Finally, the flow of document files grouped in chronological order is visualized by the arrow.

## USAGE

![image](https://user-images.githubusercontent.com/17299107/76701587-e56f5e80-6705-11ea-9005-8dcd3f3057f5.png)

**Getting started**

	python grouping_document.py
  
**Selecting directory path**

	Type directory path: "Path" ex) ./sample or C:/Users/username/Downloads/sample

**Selecting output directory path**

	Type output path: "Path" ex) ./ or d:/result
 
 **Selecting feature**
 
 	[Feature List]
	1. Company
	2. Author
	3. Last Saved By
	4. Layout
	5. RSID
	6. Company, Author, Last Saved By
	7. Everything
	
	Select the feature: number ex) 7
 
## RESULT
ex) 2020-3-15_17-35-4_Result
![results](https://user-images.githubusercontent.com/17299107/76698445-5acb3700-66e6-11ea-9203-5fb605cf21f5.png)
 
## SAMPLE DATA
'sample' foloder or http://downloads.digitalcorpora.org/corpora/files/govdocs1/by_type/docx.zip

## DEPENDENCY
- Python3.6+
- olefile
- networkx
- matplotlib
