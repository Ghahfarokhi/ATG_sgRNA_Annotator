# sgRNA_Annotator
An Excel-based tool to download the RefSeq genbank file for a given Human or Mouse gene and annotate CRISPR sgRNAs from literature. 

 sgRNA_Annotator
An Excel-based tool to download the RefSeq genbank file for a given Human or Mouse gene and annotate CRISPR sgRNAs from literature. 

•	Genome wide sgRNA libraries designed for gene knockout were collected from studies listed below and organized in a database. 
•	A Macro containing Excel workbook developed to search the database for sgRNAs designed for a given Human or Mouse gene.
•	The Excel codes are linked to control buttons in the spreadsheet graphical interface, so users will need no programming skills.
•	A user can either annotate sgRNAs on his/her genbank file, or directly use the button to download the RefSeq from NCBI and annotate sgRNAs.
•	Excel codes can recognize sgRNA pairs and calculate the distance between two cut sites. Possible dual-sgRNAs with user-specified distances will be shortlisted.
•	Deletions resulting from combination of dual-sgRNAs are listed as inFrame / FrameShift.
•	A button is provided for update purposes, then a user can occasionally check for the updates.
•	In order to use it, the library folder and the excel file must be in the same directory.

References:

Genome-wide CRISPR-guide RNA libraries (Human Genes):

GeCKOv2 library
Sanjana N. et al. (2014) Improved vectors and genome-wide libraries for CRISPR screening. Nature Methods

TKO (Toronto KnockOut) library
Hart T. et al. (2015) High-Resolution CRISPR Screens Reveal Fitness Genes and Genotype-Specific Cancer Liabilities. Cell

Knockout library
Wang T. et al. (2014) Genetic Screens in Human Cells Using the CRISPR-Cas9 System. Science

Genome-wide CRISPR-guide RNA libraries (Mouse Genes)

GeCKOv2 library
Sanjana N. et al. (2014) Improved vectors and genome-wide libraries for CRISPR screening. Nature Methods

Sanger Institute library
Koike-Yusa H. et al. (2014) Genome-wide recessive genetic screening in mammalian cells with a lentiviral CRISPR-guide RNA library. Nat Biotech

Knockout library
Schmid-Burgk J. et al. (2015) A genome-wide CRISPR screen identifies NEK7 as an essential component of NLRP3 inflammasome activation. JBC

 
Step-by-step guideline:
1-	Download and save Library folder and the sgRNA_Annotator.xlsm file in the same folder:
 
2-	Make sure that all txt files are in the Library folder:
 
3-	Open the sgRNA_Annotator.xlsm file and Enable the contents (click on the “Enable Contents” on the yellow bar appeared below the menu bar):
 
4-	Provide a gene symbol, select between Human and Mouse, and click on the “Import” button:
 
5-	If you already have the genebank file, click on the “Browse” button and select the file for annotations. Otherwise, click on “Download RefSeq file” button, and wait! The latter should download the genbank file for the RefSeq, and annotate the sgRNAs (can be found in the same folder as the Excel file is).
 
6-	Check the dual-sgRNA combination, adjust the distance between cut_sites if required and click on the “Update Dual-sgRNA” button for your changes to be applied:
 
7-	Import the generated genbank file into CLC, the sgRNAs are annotated as “CRISPR” by default.
 
Enjoy it :)

Contact info:
Amir.Taheri-Ghahfarokhi@AstraZeneca.com
PGE-Team, Discovery Sciences
