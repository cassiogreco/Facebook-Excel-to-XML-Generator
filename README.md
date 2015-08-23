# Facebook-Excel-to-XML-Generator
This is a python script that converts an excel spread sheet (xlsx) and converts it into a valid business manager-valid xml file to be uploaded as a product catalog and used for Dynamic Product Ads campaigns.

The excel sheet must have its columns and lines correctly formatted according to Facebook's documentation regarding product catalog: https://developers.facebook.com/docs/marketing-api/dynamic-product-ads/product-catalog

# Running the script
To run the script run this command: 
```python
python main.py new_generated_xml_filename.xml my_excel_file.xlsx
```

Where: ``` new_generated_xml_filename.xml ``` is the target xml file to be created, and ``` my_excel_file.xlsx ``` is your excel file
