import zipfile
import json
import logging
import os.path
import glob
import time
import openpyxl
import re


from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation

logging.basicConfig(level=os.environ.get("LOGLEVEL", "INFO"))

# ------------------------------------------------------------------------------------------------------------------------------
    # Create a new workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
# ------------------------------------------------------------------------------------------------------------------------------

# Initialising fields

Customer__Address__line1 = ""	
Customer__Address__line2 = ""	
Customer__Email = ""	
Customer__Name = ""	
Customer__PhoneNumber = ""	

Invoice__BillDetails__Name = ""	
Invoice__BillDetails__Quantity = ""	
Invoice__BillDetails__Rate = ""	
Invoice__Description = ""	
Invoice__DueDate = ""	
Invoice__IssueDate = ""	
Invoice__Number = ""	
Invoice__Tax = ""

Bussiness__City = ""	
Bussiness__Country = ""	
Bussiness__Description = ""	
Bussiness__Name = ""	
Bussiness__StreetAddress = ""	
Bussiness__Zipcode = ""	

table_heading=["Bussiness__City",	"Bussiness__Country",	"Bussiness__Description",	"Bussiness__Name",	"Bussiness__StreetAddress",	"Bussiness__Zipcode",	"Customer__Address__line1",	"Customer__Address__line2",	"Customer__Email",	"Customer__Name",	"Customer__PhoneNumber",	"Invoice__BillDetails__Name",	"Invoice__BillDetails__Quantity",	"Invoice__BillDetails__Rate",	"Invoice__Description",	"Invoice__DueDate",	"Invoice__IssueDate",	"Invoice__Number",	"Invoice__Tax"]
sheet.append(table_heading)
row_data=[Bussiness__City,	Bussiness__Country,	Bussiness__Description,	Bussiness__Name,	Bussiness__StreetAddress,	Bussiness__Zipcode,	Customer__Address__line1,	Customer__Address__line2,	Customer__Email,	Customer__Name,	Customer__PhoneNumber,	Invoice__BillDetails__Name,	Invoice__BillDetails__Quantity,	Invoice__BillDetails__Rate,	Invoice__Description,	Invoice__DueDate,	Invoice__IssueDate,	Invoice__Number,	Invoice__Tax]

# required flags
date_count=0
data_initialised=False
details_left_bound=240.25999450683594
payment_left_bound=412.8000030517578
col_num=1

# regex for info extraction
date_pattern = r"\b(\d{2}-\d{2}-\d{4})\b"
invoice_pattern = r'Invoice#\s*([A-Za-z0-9]+)'
mobile_number_pattern = r"\b(\d{3}-\d{3}-\d{4})\b"

try:
    # get base path.
    base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Initial setup, create credentials instance.
    credentials = Credentials.service_principal_credentials_builder(). \
        with_client_id(os.getenv('PDF_SERVICES_CLIENT_ID')). \
        with_client_secret(os.getenv('PDF_SERVICES_CLIENT_SECRET')). \
        build()

    # Create an ExecutionContext using credentials and create a new operation instance.
    execution_context = ExecutionContext.create(credentials)
    extract_pdf_operation = ExtractPDFOperation.create_new()

    # Find all PDF files in the resources folder and sort numerically not lexicographically
    pdf_collection = sorted(glob.glob(base_path + "/resources/output*.pdf"), key=lambda filename: len(filename))


    for pdf_file in pdf_collection:
        # reinitialising for each pdf
        date_count=0
        Invoice__Description = ""
        # Set operation input from the current PDF file.
        source = FileRef.create_from_local_file(pdf_file)
        extract_pdf_operation.set_input(source)

        # Build ExtractPDF options and set them into the operation
        extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
            .with_element_to_extract(ExtractElementType.TEXT) \
            .build()
            # .with_element_to_extract(ExtractElementType.TABLES) \
        extract_pdf_operation.set_options(extract_pdf_options)

        # Execute the operation.
        result: FileRef = extract_pdf_operation.execute(execution_context)

        output_file = os.path.splitext(os.path.basename(pdf_file))[0] + "_" + str(int(time.time())) + ".zip"
        output_path = os.path.join(base_path, "secondary_generated_resources", output_file)
        result.save_as(output_path)

        # Extract and print data from the structuredData.json file within the ZIP archive
        with zipfile.ZipFile(output_path, 'r') as archive:
            json_entry = archive.open('structuredData.json')
            json_data = json_entry.read()
            data = json.loads(json_data)
            text_list = []
            for element in data['elements']:
                if 'Text' in element:

                    # getting bounds
                    left_bound=element['Bounds'][0]

                    text_list.append(element['Text'])
                    temp_text=element['Text']

                    # ------getting Invoice Details--------
                    if(left_bound==details_left_bound):
                        if (temp_text.strip()!="DETAILS"):
                            Invoice__Description+=temp_text
                        
                            
                    # ----------getting dates--------------
                    match=re.search(date_pattern,temp_text)
                    if match:
                        date = match.group(1)
                        date_count+=1
                        if(date_count==1):
                            Invoice__IssueDate=date
                        elif(date_count==2):
                            Invoice__DueDate=date

                    # -------getting invoice number--------
                    match = re.search(invoice_pattern, temp_text)
                    if match:
                        Invoice__Number = match.group(1)

                    # -------getting mobile number--------
                    match = re.search(mobile_number_pattern, temp_text)
                    if match:
                        Customer__PhoneNumber = match.group(1)
                        

            if not(data_initialised):
                data_initialised=True

                preprocessed_address=text_list[1].split(',')

                Bussiness__StreetAddress=preprocessed_address[0].strip()
                Bussiness__City=preprocessed_address[1].strip()

                Bussiness__Name=text_list[0].strip()
                Bussiness__Country=text_list[2].strip()
                Bussiness__Zipcode=text_list[3].strip()
                Bussiness__Description=text_list[8].strip()
            
            for i in range(len(text_list)):
                if(text_list[i].strip()=="AMOUNT"):
                    i+=1
                    while(text_list[i].strip()!="Subtotal"):
                        Invoice__BillDetails__Name=text_list[i+0].strip()
                        Invoice__BillDetails__Quantity=int(text_list[i+1].strip())
                        Invoice__BillDetails__Rate=int(text_list[i+2].strip())
                        i+=4
                        
                        Invoice__Tax=int(text_list[-3].strip())
                        row_data=[Bussiness__City,	Bussiness__Country,	Bussiness__Description,	Bussiness__Name,	Bussiness__StreetAddress,	int(Bussiness__Zipcode),	Customer__Address__line1,	Customer__Address__line2,	Customer__Email,	Customer__Name,	Customer__PhoneNumber,	Invoice__BillDetails__Name,	Invoice__BillDetails__Quantity,	Invoice__BillDetails__Rate,	Invoice__Description.strip(),	Invoice__DueDate,	Invoice__IssueDate,	Invoice__Number,	Invoice__Tax]
                        # print(row_data)
                        sheet.append(row_data)


            # for index, text in enumerate(text_list, start=1):
            #     sheet.cell(row=index, column=col_num).value = text
            # row_data=[Bussiness__City,	Bussiness__Country,	Bussiness__Description,	Bussiness__Name,	Bussiness__StreetAddress,	int(Bussiness__Zipcode),	Customer__Address__line1,	Customer__Address__line2,	Customer__Email,	Customer__Name,	Customer__PhoneNumber,	Invoice__BillDetails__Name,	Invoice__BillDetails__Quantity,	Invoice__BillDetails__Rate,	Invoice__Description.strip(),	Invoice__DueDate,	Invoice__IssueDate,	Invoice__Number,	Invoice__Tax]
            # sheet.append(row_data)

        col_num+=1

# ------------------------------------------------------------------------------------------------------------------------------
    output_excel_file = str(int(time.time()))+"_ExtractedData.xlsx" 
    excel_output_path = os.path.join(base_path, "output",output_excel_file)
    workbook.save(excel_output_path)
    # Close the workbook
    workbook.close()
# ------------------------------------------------------------------------------------------------------------------------------

except (ServiceApiException, ServiceUsageException, SdkException):
    logging.exception("Exception encountered while executing operation")
