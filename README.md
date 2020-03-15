# Grouping_document_tool

## INTRODUCTION

The author developed an automation tool for document grouping for the experiment. Developed tools can be grouped by selecting the desired feature (Company, author, last saved by, layout, RSID, property information, and everything). Considering the efficiency of the investigation, the tool can easily view only the grouped document files by copying the grouped document files and storing them in folders. In addition, the details of the grouped document files (filename, company, author, last modified by, etc.) are saved in the .csv file, as shown in Table.~\ref{tab:metadataList}. Finally, the flow of document files grouped in chronological order, such as Fig.~\ref{fig:visualization_final}, is visualized by the arrow. We uploaded developed tool on Github.

## OVERVIEW

https://github.com/blackmax90/Grouping_document_tool/issues/1#issue-581583248

## USAGE

**Getting started**
	python grouping_document.py
  
**Selecting directory path**

	Type directory path: "Path" ex) ./sample or sample
 
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
 
## SAMPLE DATA
'sample' foloder or http://downloads.digitalcorpora.org/corpora/files/govdocs1/by_type/docx.zip

## DEPENDENCY
