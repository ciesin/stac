# CIESIN STAC and STAC Browser - Alpha

This repository houses all of the related files to building the CIESIN STAC. 

These files include excel workbooks that represent Items and Assets, scripts, a stac system including folders and json files and a stac browser comprised of index.html, app.js, and style.css files. 



## CIESIN STAC workflow readme

The CIESIN STAC (SpatioTemporal Asset Catalog) is built using a combination of python scripts and excel workbooks that use Harvard Dataverse, Columbia Academic Commons, NASA CMR (Earthdata), and existing local XML files. There are two main scripts that build the two main Excel workbooks, which refer to the other scripts as packages, therefore, the scripts in the folder should not be erased, edited, or moved without understanding where there are necessary changes in other scripts as a result. 
The difference between the two main scripts is that one of the main scripts builds the STAC on existing Items workbook, updating missing cells and finding Assets. The other uses a list of DOIs to do the same as the first, as well as finding potential new Items as well. Both workflows output two updated excel workbooks, representing STAC Items and related STAC Assets, respectively. 
The primary script is:
build_ciesin_items_and_assets_workbooks.py
Builds the initial Items and Assets Excel workbooks from your main inventory.
It contains the core functionality for retrieving, extracting, normalizing, and assembling dataset metadata and assets from multiple sources.
The companion script:
update_ciesin_items_workbook.py
Takes previously generated workbooks, identifies incomplete records, fills missing metadata, and updates the Assets workbook while preserving existing information, discovers new assets, and generates updated Items and Assets workbooks. 

The most important distinction is that the first script is a workbook generation process, while the second is an incremental metadata enrichment process.
Once the Items and Assets workbook are approved, a third script develops the final STAC (folders and json files):
build_stac_from_ciesin_workbooks.py


#### Main automations
Extract metadata from ArcGIS ISO 19139 XML files.
Retrieve metadata using dataset DOIs and external metadata services.
Combine metadata from multiple sources into consistent dataset records.
Identify and associate dataset assets with their parent datasets.
Normalize and classify asset information.
Identify duplicate assets while preserving valid dataset-to-asset relationships.
Maintain existing inventory values and fill missing metadata.
Generate Excel workbooks for the Items and Assets inventories. 
Inputs and directory structure
The update workflow uses three main inputs:
An existing CIESIN Items workbook.
An optional paired Assets workbook.
A directory containing all ArcGIS XML metadata records available.


##### The project structure is:

project/
│
├── scripts/
│   ├── build_ciesin_items_and_assets_workbooks.py
│   ├── build_stac_from_ciesin_workbooks.py
│   ├── update_ciesin_items_workbook.py
│   ├── OTHER SCRIPTS
│
├── inputs/
│   └── xmls/
│       ├── dataset_1.xml
│       ├── dataset_2.xml
│       └── ...
│
└── outputs/
    ├── stac_from_workbooks/
    │	└── stac/  
    └── stac_workbooks/
        ├── ciesin_items_*.xlsx
        └── ciesin_assets_*.xlsx

#### Items workbook
The input workbook must contain an Item worksheet (or the worksheet specified with --sheet).
The following columns are required by the update script:
item_title
DOI columns for Academic Commons, Harvard Dataverse, and NASA CMR. The DOIs can be blank, but at least one DOI is required. 
inventory_xml_location
The builder uses the available identifiers and metadata to locate and enrich each dataset.
#### Assets workbook
The Assets workbook contains existing asset records associated with the Items workbook.
When --assets-input is not specified, the update script attempts to infer the paired Assets workbook from the Items filename.
For example:
ciesin_items_20261020T120000Z.xlsx
ciesin_assets_20261020T120000Z.xlsx 
If no paired Assets workbook is found, the workflow starts without existing asset records.  
XML metadata directory
The default XML directory is:
inputs/xmls/
This can be overridden with --xml-dir.
XML records provide an additional source of metadata, particularly for datasets with ArcGIS ISO 19139 metadata and the FGDC standards.  

#### Metadata extraction and enrichment

The builder combines information from the existing inventory, XML metadata, and external metadata sources.
Depending on the available source information, metadata fields can include:
Dataset title
Resource type
STAC category
Tags and keywords
Spatial extent and bounding box
Temporal coverage
License
Description
Providers
Citation
Version
Additional resources
Purpose
Authors
Publishers
Place
DOI
Publication date
Importantly, candidate metadata does not automatically overwrite the existing workbook. The scripts fill blank cells.
DOI lookup
DOIs are used to identify and retrieve additional dataset metadata. 
Multiple DOI fields can be associated with a single dataset. The workflow uses the DOI columns to identify and retrieve relevant metadata.

#### Processing sequence
The update script follows this general sequence:
Read the input Items workbook.
Identify the requested worksheet and validate required columns.
Locate and read the paired Assets workbook, if available.
Match existing assets to their parent datasets using dataset_id or DOI.
Check metadata completeness for each dataset.
Skip DOI lookups for complete records.
Retrieve metadata for incomplete records for online sources (HDV, NASA, AC), unless offline mode is enabled.
Generate candidate metadata.
Fill blank cells in the existing Items worksheet.
Update the dataset context with final identifiers and values.
Discover Assets related to the Items.
Merge existing and newly discovered Assets.
Remove duplicate Harvard Academic Commons record-page assets when a corresponding DOI asset exists.
Classify missing asset subtypes.
Generate the updated Items and Assets workbooks.
Add update logs and field-level update details.
Preserving existing metadata
Existing populated values are retained.
Missing values can be filled using candidate metadata.
Fields without a candidate value remain blank.
Existing records are not replaced wholesale by external metadata.

##### Offline mode

The --offline option disables remote DOI lookups:
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_example.xlsx `
  --offline
When offline mode is enabled:
Remote DOI lookups are skipped.
XML metadata can still be used.
Candidate metadata is assembled from the available local information.
Existing workbook values remain protected.
Offline mode is useful for testing XML extraction, updating records without accessing external APIs, or extracting asset information from the XML files.  

Matching assets to datasets
Existing asset records are matched to inventory rows using:
dataset_id, when available. DOI identifiers, when a dataset ID match is not available. 
 
### Running the workflow
Run commands from the project root. For example 
cd C:\STAC_workflow
This folder should contain the script folder and the excel workboks 
Basic update
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx
The script automatically:
Uses the Items worksheet.
Uses inputs/xmls/ as the XML directory.
Looks for the paired Assets workbook.
Writes output files to the input workbook's directory.
Uses a default pause of 1.5 seconds between DOI lookup operations.
Specify the Assets workbook
If the paired Assets workbook cannot be inferred automatically:
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx `
  --assets-input outputs/stac_workbooks/ciesin_assets_20261020T120000Z.xlsx
Specify the XML directory
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx `
  --xml-dir inputs/xmls
Specify an output directory
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx `
  --output-dir outputs/stac_workbooks
Change the worksheet
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx `
  --sheet Items
Adjust the pause between lookups
python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx `
  --pause-seconds 2.0

#### Run offline

python scripts/update_ciesin_items_workbook.py `
  --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx `
  --offline
Available arguments
Argument
Description
Default
--input
Previously generated Items workbook. Required.
None
--assets-input
Existing paired Assets workbook.
Inferred from Items filename
--sheet
Input worksheet name.
Items
--xml-dir
Directory containing XML metadata.
inputs/xmls/
--output-dir
Output directory.
Input workbook directory
--pause-seconds
Delay between DOI lookup operations.
1.5
--offline
Skip remote DOI lookups.
Disabled

Use --help to see the current command-line options:
python scripts/update_ciesin_items_workbook.py --help

#### Output workbooks
The update script generates two timestamped workbooks. One for Items and onfe for Assets. The workbook contains: Items worksheet. The updated dataset inventory. Existing populated cells are preserved, and blank cells are filled when candidate metadata is available.
Existing values are authoritative
The update script is designed to fill missing metadata rather than overwrite existing values.
If an existing value is incorrect, it should be reviewed and corrected deliberately rather than expecting the update workflow to replace it automatically.
If a search is needed to correct the cell, delete the existing cell contents and run the scripts. 
THINGS TO REMEMBER
Complete records are skipped for DOI lookups. Records considered complete are excluded from remote DOI searches. If additional assets or metadata need to be discovered for an otherwise complete dataset, the completeness logic may need to be adjusted or the record's missing fields addressed deliberately.
Assets and metadata have separate workflows. Dataset metadata is assembled through make_items(), while asset discovery is handled through assemble_assets(). Changes to asset extraction should generally be made in the builder's asset-related functions rather than in the workbook update logic.
Asset deduplication must preserve dataset relationships. Do not globally deduplicate assets based only on their URL. A single URL may be associated with multiple datasets. The intended duplicate check is based on the dataset identity and asset URL.
External metadata sources can fail independently. The workflow relies on external metadata services and local XML records. When a source is unavailable or does not return a match, the builder should retain other available metadata rather than discarding the dataset record. Rerun the script if the record does not return any new metadata. This issue is particularly true with Columbia Academic Commons. 
#### Troubleshooting
Input workbook not found
Check the --input path and confirm that the workbook exists.
Required columns missing
The update script requires:
item_title
The DOI columns defined in the builder
inventory_xml_location
Check the worksheet headers and ensure they match the expected names.
Existing Assets workbook not found
Check the paired filename or explicitly provide the workbook using --assets-input.
Without an existing Assets workbook, the script cannot use previous asset records when checking completeness or merging discovered assets.
No metadata fields are filled
Check:
Whether the dataset is already considered complete.
Whether the DOI columns contain valid identifiers.
Whether XML metadata is available.
Whether remote DOI lookups returned matching records.
Whether the candidate metadata contains values for the missing fields.
The Update details worksheet provides the candidate and final values for each field.
#### Recommended development workflow
When modifying the project:
Identify whether the change belongs in the core builder or the update wrapper.
Make metadata extraction and source lookup changes in build_ciesin_items_and_assets_workbooks.py.
Keep workbook orchestration, completeness checks, merging, and update logging in update_ciesin_items_workbook.py.
Test changes against multiple datasets and metadata sources.
Check both the Items and Assets outputs.
Review the Update details and Update log worksheets.
Confirm that existing populated values and valid dataset-to-asset relationships are preserved.
For source-specific changes, test datasets that exercise different metadata structures and lookup paths rather than relying on a single example.

### Workflow summary
The relationship between the two main scripts is:
Existing Items workbook
        |
        v
update_ciesin_items_workbook.py
        |
        +---- Read existing Items and Assets
        |
        +---- Check metadata completeness
        |
        +---- Identify incomplete datasets
        |
        v
build_ciesin_items_and_assets_workbooks.py
        |
        +---- Retrieve DOI metadata
        |
        +---- Extract XML metadata
        |
        +---- Assemble candidate Items
        |
        +---- Discover dataset assets
        |
        v
Update existing Items workbook
        |
        +---- Fill blank metadata fields
        |
        +---- Merge existing and discovered assets
        |
        +---- Deduplicate and classify assets
        |
        v
Updated Items workbook
Updated Assets workbook
        |
        +---- Update details
        +---- Update log


### Building STAC from Workbooks.

Once the Items and Assets excel workbooks are finalized, the workbooks can be used to build a CIESIN STAC architecture. This workflow would produce a ‘stac’ folder with subfolders and json files that construct the STAC.

To build from the two excel workbooks, use the script 

build_stac_from_ciesin_workbooks.py

The script requires --items and --assets variables pointing to the excel workbooks to be used. It also requires --output-dir for the output directory of the stac folder.

python .\scripts\build_stac_from_ciesin_workbooks.py --items .\inputs\stac_workbooks\ciesin_items_v2.xlsx --assets .\inputs\stac_workbooks\ciesin_assets_v2.xlsx  --output-dir .\outputs\stac_from_workbooks


