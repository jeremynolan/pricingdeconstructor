from flask import Flask, request, make_response
import pandas as pd
import plotly.express as px
import plotly.io as pio
import os
import logging
from datetime import datetime
import openpyxl
import re
import string
import secrets
# import boto3  # Uncomment for S3 support

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', secrets.token_hex(16))
UPLOAD_FOLDER = '/tmp/Uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Logging: WARNING to minimize L11 errors, INFO for key ops
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Pricing rules
pricing_rules = {
    "Process": {},
    "Coating": {},
    "Foil Material": {},
    "Foil Thickness": {},
    "Colour": {}
}

# Process and Step Process mapping
process_step_mapping = {
    "Chemetch": ["Single", "Double", "Triple", "5 or more"],
    "LaserSTEP": ["1-2", "1-5", "1-10", "1-15", "1-20", "21-30", "31-40", "41-50", "51-60"],
    "Milled": ["Single", "Double", "Triple", "Quad"],
    "LaserCut": []
}

# Inline CSS
css = """
<style>
    body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f9; }
    .container { max-width: 1200px; margin: 0 auto; }
    h1 { color: #333; text-align: center; }
    .form-group { margin-bottom: 15px; }
    label { display: inline-block; width: 200px; font-weight: bold; }
    input[type="file"], input[type="number"] { padding: 5px; width: 200px; }
    button { padding: 10px 20px; background-color: #007bff; color: white; border: none; cursor: pointer; }
    button:hover { background-color: #0056b3; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #007bff; color: white; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .error { color: red; text-align: center; }
    .download { display: inline-block; margin-top: 20px; margin-right: 10px; padding: 10px 20px; background-color: #28a745; color: white; text-decoration: none; }
    .download:hover { background-color: #218838; }
    .download-excel { background-color: #17a2b8; }
    .download-excel:hover { background-color: #138496; }
    #chart { margin-top: 20px; }
</style>
"""

# Upload page HTML
upload_html = """
<!DOCTYPE html>
<html>
<head><title>Price Deconstructor</title>
""" + css + """
</head>
<body>
    <div class="container">
        <h1>Upload Sales Report</h1>
        <p>Ensure your Excel file has the sheet "SalesbyItemBASEPRICEDECON" with columns: Sales Price, Frame, Customer/Project: Company Name, Process, Step Process, Coating, Foil Material, Foil Thickness, Colour.</p>
        <p>Pricing file (optional) format: <code>chem Single: 100</code>, <code>laserstep 1-2: 200</code>, <code>coat Advanced Nano: 50</code>.</p>
        {{error|safe}}
        <form method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label for="excel_file">Select Excel File (.xlsx):</label>
                <input type="file" id="excel_file" name="excel_file" accept=".xlsx" required>
            </div>
            <div class="form-group">
                <label for="pricing_file">Select Pricing File (.txt, optional):</label>
                <input type="file" id="pricing_file" name="pricing_file" accept=".txt">
            </div>
            <button type="submit">Upload & Proceed to Pricing</button>
        </form>
    </div>
</body>
</html>
"""

# Pricing form HTML
pricing_form_html = """
<!DOCTYPE html>
<html>
<head><title>Pricing Rules</title>
""" + css + """
</head>
<body>
    <div class="container">
        <h1>Enter Pricing Rules</h1>
        <p>Enter prices manually or use uploaded pricing file values. Re-upload Excel file to ensure processing.</p>
        {{error|safe}}
        <form method="post" action="/pricing" enctype="multipart/form-data">
            <div class="form-group">
                <label for="excel_file">Re-upload Excel File (.xlsx):</label>
                <input type="file" id="excel_file" name="excel_file" accept=".xlsx" required>
            </div>
            <h3>Process and Step Process</h3>
            {% for process in processes %}
            <div class="form-group">
                <h4>{{process}}</h4>
                {% for step in process_step_mapping[process] %}
                <div class="form-group">
                    <label for="{{process}}_{{step}}">{{step}}</label>
                    <input type="number" step="0.01" id="{{process}}_{{step}}" name="{{process}_{{step}}" placeholder="Cost ($)" value="{{form_data.get(process ~ '_' ~ step, '')}}">
                </div>
                {% endfor %}
            </div>
            {% endfor %}
            <h3>Coating</h3>
            {% for coating in ['Advanced Nano', 'Nano Wipe', 'Nano Slic', 'BluPrint'] %}
            <div class="form-group">
                <label for="Coating_{{coating}}">{{coating}}</label>
                <input type="number" step="0.01" id="Coating_{{coating}}" name="Coating_{{coating}}" placeholder="Cost ($)" value="{{form_data.get('Coating_' ~ coating, '')}}">
            </div>
            {% endfor %}
            <button type="submit">Process File</button>
        </form>
    </div>
</body>
</html>
"""

# Results page HTML
results_html = """
<!DOCTYPE html>
<html>
<head>
    <title>Results</title>
""" + css + """
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
</head>
<body>
    <div class="container">
        <h1>Deconstructed Pricing</h1>
        <p>{{data|length}} Unique Customer-Material Combinations Processed</p>
        {{error|safe}}
        <a href="/download" class="download">Download Results as CSV</a>
        <a href="/download_excel" class="download download-excel">Download Results as Excel</a>
        <h3>Lowest Base Cost by Customer</h3>
        <div id="chart">{{chart|safe}}</div>
        <table>
            <thead>
                <tr>
                    <th>Customer</th>
                    <th>Frame</th>
                    <th>Sales Price</th>
                    <th>Process</th>
                    <th>Step Process</th>
                    <th>Coating</th>
                    <th>Foil Material</th>
                    <th>Foil Thickness</th>
                    <th>Colour</th>
                    <th>Attribute Cost</th>
                    <th>Base Cost</th>
                </tr>
            </thead>
            <tbody>
                {% for row in data %}
                <tr>
                    <td>{{row.Customer}}</td>
                    <td>{{row.Frame}}</td>
                    <td>{{row.Sales_Price}}</td>
                    <td>{{row.Process}}</td>
                    <td>{{row.Step_Process}}</td>
                    <td>{{row.Coating}}</td>
                    <td>{{row.Foil_Material}}</td>
                    <td>{{row.Foil_Thickness}}</td>
                    <td>{{row.Colour}}</td>
                    <td>{{row.Attribute_Cost}}</td>
                    <td>{{row.Base_Cost}}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

def sanitize_filename(filename):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    sanitized = ''.join(c if c in valid_chars else '_' for c in filename)
    sanitized = re.sub(r'_+', '_', sanitized)
    return sanitized.strip('_')

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    logger.info("Entering / route")
    if request.method == 'POST':
        logger.info("Received POST request for file upload")
        try:
            excel_file = request.files.get('excel_file')
            pricing_file = request.files.get('pricing_file')
            if not excel_file:
                logger.error("No Excel file provided")
                return app.jinja_env.from_string(upload_html).render(error='<p class="error">No Excel file selected. Please choose a file.</p>')
            
            if not excel_file.filename.endswith('.xlsx'):
                logger.error(f"Invalid Excel file extension: {excel_file.filename}")
                return app.jinja_env.from_string(upload_html).render(error='<p class="error">Please upload a valid .xlsx file.</p>')
            
            # Save Excel file
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            sanitized_filename = sanitize_filename(excel_file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, f"{timestamp}_{sanitized_filename}")
            file_path = os.path.normpath(file_path)
            logger.info(f"Saving Excel file to: {file_path}")
            
            if not os.access(UPLOAD_FOLDER, os.W_OK):
                logger.error(f"No write permissions for Uploads folder: {UPLOAD_FOLDER}")
                return app.jinja_env.from_string(upload_html).render(error='<p class="error">Server error: No write permissions for Uploads folder.</p>')
            
            excel_file.save(file_path)
            # Uncomment for S3
            # s3 = boto3.client('s3', aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'), aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'))
            # s3.upload_file(file_path, 'your-bucket', os.path.basename(file_path))
            
            if not os.path.exists(file_path):
                logger.error(f"Excel file not found after saving: {file_path}")
                return app.jinja_env.from_string(upload_html).render(error='<p class="error">Failed to save file: {excel_file.filename}. Please try again.</p>')
            
            # Validate Excel file
            logger.info(f"Validating Excel file: {file_path}")
            wb = openpyxl.load_workbook(file_path)
            sheet_names = wb.sheetnames
            if 'SalesbyItemBASEPRICEDECON' not in sheet_names:
                logger.error(f"Sheet 'SalesbyItemBASEPRICEDECON' not found")
                try:
                    os.remove(file_path)
                    logger.info(f"Removed invalid file: {file_path}")
                except Exception as e:
                    logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
                return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Sheet "SalesbyItemBASEPRICEDECON" not found in {excel_file.filename}. Available sheets: {", ".join(sheet_names)}</p>')
            
            df = pd.read_excel(file_path, sheet_name='SalesbyItemBASEPRICEDECON', engine='openpyxl', nrows=1)
            actual_columns = [str(col).strip().lower() for col in df.columns]
            required_columns = ['Sales Price', 'Frame', 'Customer/Project: Company Name', 'Process', 'Step Process', 'Coating', 'Foil Material', 'Foil Thickness', 'Colour']
            required_columns_normalized = [col.strip().lower() for col in required_columns]
            missing_columns = [col for col in required_columns if col.strip().lower() not in actual_columns]
            if missing_columns:
                logger.error(f"Missing columns in Excel file: {missing_columns}")
                try:
                    os.remove(file_path)
                    logger.info(f"Removed invalid file: {file_path}")
                except Exception as e:
                    logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
                return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Missing columns in {excel_file.filename}: {', '.join(missing_columns)}. Found: {', '.join(df.columns)}</p>')
            
            logger.info("Excel file validated successfully")
            
            # Handle pricing file if provided
            form_data = {}
            if pricing_file and pricing_file.filename.endswith('.txt'):
                try:
                    logger.info(f"Processing pricing file: {pricing_file.filename}")
                    pricing_file_path = os.path.join(UPLOAD_FOLDER, f"{timestamp}_{sanitize_filename(pricing_file.filename)}")
                    pricing_file_path = os.path.normpath(pricing_file_path)
                    pricing_file.save(pricing_file_path)
                    
                    if not os.path.exists(pricing_file_path):
                        logger.error(f"Pricing file not found after saving: {pricing_file_path}")
                        return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Failed to save pricing file: {pricing_file.filename}. Please try again.</p>')
                    
                    with open(pricing_file_path, 'r') as f:
                        for line in f:
                            line = line.strip()
                            if not re.match(r'^\w+\s*.*?:\s*\d+(\.\d+)?$', line):
                                continue
                            key, value = [part.strip() for part in line.split(':', 1)]
                            try:
                                value = float(value)
                            except ValueError:
                                continue
                            
                            key = key.lower().replace('lasterstep', 'laserstep').replace('laststep', 'laserstep')
                            if key.startswith('chem '):
                                step = key[5:]
                                if step == '5 or more' or step.title() in process_step_mapping["Chemetch"]:
                                    form_data[f"Chemetch_{step if step == '5 or more' else step.title()}"] = str(value)
                            elif key.startswith('laserstep '):
                                step = key[10:]
                                if step in process_step_mapping["LaserSTEP"]:
                                    form_data[f"LaserSTEP_{step}"] = str(value)
                                elif step in ["21-30", "31-40", "41-50", "51-60"]:
                                    form_data[f"LaserSTEP_{step}"] = str(245)
                            elif key.startswith('mill '):
                                step = key[5:].title()
                                if step in process_step_mapping["Milled"]:
                                    form_data[f"Milled_{step}"] = str(value)
                            elif key == 'double':
                                form_data["Milled_Double"] = str(value)
                            elif key.startswith('coat '):
                                coating = key[5:].title().replace('Bluprint', 'BluPrint')
                                if coating in ["Advanced Nano", "Nano Wipe", "Nano Slic", "BluPrint"]:
                                    form_data[f"Coating_{coating}"] = str(value)
                    
                    try:
                        os.remove(pricing_file_path)
                        logger.info(f"Removed temporary pricing file: {pricing_file_path}")
                    except Exception as e:
                        logger.warning(f"Failed to remove pricing file {pricing_file_path}: {str(e)}")
                
                except Exception as e:
                    logger.error(f"Error processing pricing file: {str(e)}")
                    return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Error processing pricing file: {pricing_file.filename}. Please try again.</p>')
            
            try:
                os.remove(file_path)
                logger.info(f"Removed Excel file after validation: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to remove Excel file {file_path}: {str(e)}")
            
            return app.jinja_env.from_string(pricing_form_html).render(
                processes=process_step_mapping.keys(),
                process_step_mapping=process_step_mapping,
                form_data=form_data,
                error=None
            )
        except Exception as e:
            logger.error(f"Unexpected error during file upload: {str(e)}")
            return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Unexpected error during upload: {str(e)}. Please try again.</p>')
    
    logger.info("Rendering upload page for GET request")
    return app.jinja_env.from_string(upload_html).render(error=None)

@app.route('/pricing', methods=['GET', 'POST'])
def pricing_form():
    logger.info("Entering /pricing route")
    if request.method == 'GET':
        logger.info("Accessed /pricing via GET, redirecting to upload")
        return app.jinja_env.from_string(upload_html).render(error='<p class="error">Please upload an Excel file first.</p>')
    
    logger.info("Processing pricing form submission")
    global pricing_rules
    pricing_rules = {
        "Process": {},
        "Coating": {},
        "Foil Material": {},
        "Foil Thickness": {},
        "Colour": {}
    }
    
    try:
        excel_file = request.files.get('excel_file')
        if not excel_file:
            logger.error("No Excel file re-uploaded in pricing form")
            return app.jinja_env.from_string(pricing_form_html).render(
                processes=process_step_mapping.keys(),
                process_step_mapping=process_step_mapping,
                form_data={},
                error='<p class="error">Please re-upload the Excel file.</p>'
            )
        
        if not excel_file.filename.endswith('.xlsx'):
            logger.error(f"Invalid Excel file extension: {excel_file.filename}")
            return app.jinja_env.from_string(pricing_form_html).render(
                processes=process_step_mapping.keys(),
                process_step_mapping=process_step_mapping,
                form_data={},
                error='<p class="error">Please upload a valid .xlsx file.</p>'
            )
        
        # Save re-uploaded Excel file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        sanitized_filename = sanitize_filename(excel_file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, f"{timestamp}_{sanitized_filename}")
        file_path = os.path.normpath(file_path)
        logger.info(f"Saving re-uploaded Excel file to: {file_path}")
        
        if not os.access(UPLOAD_FOLDER, os.W_OK):
            logger.error(f"No write permissions for Uploads folder: {UPLOAD_FOLDER}")
            return app.jinja_env.from_string(pricing_form_html).render(
                processes=process_step_mapping.keys(),
                process_step_mapping=process_step_mapping,
                form_data={},
                error='<p class="error">Server error: No write permissions for Uploads folder.</p>'
            )
        
        excel_file.save(file_path)
        # Uncomment for S3
        # s3 = boto3.client('s3', aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'), aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'))
        # s3.upload_file(file_path, 'your-bucket', os.path.basename(file_path))
        
        if not os.path.exists(file_path):
            logger.error(f"Excel file not found after saving: {file_path}")
            return app.jinja_env.from_string(pricing_form_html).render(
                processes=process_step_mapping.keys(),
                process_step_mapping=process_step_mapping,
                form_data={},
                error='<p class="error">Failed to save Excel file: {excel_file.filename}. Please try again.</p>'
            )
        
        # Process form data
        form_data = {key: value for key, value in request.form.items()}
        non_zero_prices = False
        for process in process_step_mapping:
            pricing_rules["Process"][process] = {}
            for step in process_step_mapping[process]:
                cost = request.form.get(f"{process}_{step}", "0")
                try:
                    cost_value = float(cost) if cost.strip() else 0
                    if process == "LaserSTEP" and step in ["21-30", "31-40", "41-50", "51-60"] and cost_value == 0:
                        cost_value = pricing_rules["Process"]["LaserSTEP"].get("1-20", 245)
                    pricing_rules["Process"][process][step] = cost_value
                    if cost_value != 0:
                        non_zero_prices = True
                except ValueError:
                    pricing_rules["Process"][process][step] = 0
        
        for coating in ["Advanced Nano", "Nano Wipe", "Nano Slic", "BluPrint"]:
            cost = request.form.get(f"Coating_{coating}", "0")
            try:
                cost_value = float(cost) if cost.strip() else 0
                pricing_rules["Coating"][coating] = cost_value
                if cost_value != 0:
                    non_zero_prices = True
            except ValueError:
                pricing_rules["Coating"][coating] = 0
        
        if not non_zero_prices:
            logger.error("No non-zero pricing rules provided")
            return app.jinja_env.from_string(pricing_form_html).render(
                processes=process_step_mapping.keys(),
                process_step_mapping=process_step_mapping,
                form_data=form_data,
                error='<p class="error">Please provide at least one non-zero pricing rule.</p>'
            )
    except Exception as e:
        logger.error(f"Error processing pricing form: {str(e)}")
        return app.jinja_env.from_string(pricing_form_html).render(
            processes=process_step_mapping.keys(),
            process_step_mapping=process_step_mapping,
            form_data=form_data,
            error=f'<p class="error">Error processing pricing form: {str(e)}. Please try again.</p>'
        )
    
        # Process Excel file
    logger.info(f"Validating re-uploaded file: {file_path}")
    if not os.access(file_path, os.R_OK):
        logger.error(f"No read permissions for file: {file_path}")
        try:
            os.remove(file_path)
            logger.info(f"Removed invalid file: {file_path}")
        except Exception as e:
            logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
        return app.jinja_env.from_string(pricing_form_html).render(
            processes=process_step_mapping.keys(),
            process_step_mapping=process_step_mapping,
            form_data=form_data,
            error='<p class="error">No read permissions for file: {os.path.basename(file_path)}. Please upload again.</p>'
        )
    
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.sheetnames
        if 'SalesbyItemBASEPRICEDECON' not in sheet_names:
            logger.error(f"Sheet 'SalesbyItemBASEPRICEDECON' not found in {file_path}")
            try:
                os.remove(file_path)
                logger.info(f"Removed invalid file: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
            return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Sheet "SalesbyItemBASEPRICEDECON" not found in {os.path.basename(file_path)}. Available sheets: {", ".join(sheet_names)}</p>')
        
        # Process Excel in chunks
        results = []
        chunk_size = 5000
        for chunk in pd.read_excel(file_path, sheet_name='SalesbyItemBASEPRICEDECON', engine='openpyxl', chunksize=chunk_size):
            logger.info(f"Processing chunk with {len(chunk)} rows")
            for index, row in chunk.iterrows():
                try:
                    if pd.isna(row['Sales Price']) or pd.isna(row['Frame']) or pd.isna(row['Customer/Project: Company Name']):
                        continue
                    process = str(row['Process']).strip() if not pd.isna(row['Process']) else 'Unknown'
                    step_process = str(row['Step Process']).strip() if not pd.isna(row['Step Process']) else 'None'
                    if process == 'LaserSTEP':
                        step_process = re.sub(r'\s*-\s*', '-', step_process)
                    coating = str(row['Coating']).strip() if not pd.isna(row['Coating']) else 'None'
                    foil_material = str(row['Foil Material']).strip() if not pd.isna(row['Foil Material']) else 'Unknown'
                    foil_thickness = str(row['Foil Thickness']).strip() if not pd.isna(row['Foil Thickness']) else 'Unknown'
                    colour = str(row['Colour']).strip() if not pd.isna(row['Colour']) else 'Unknown'
                    customer = str(row['Customer/Project: Company Name']).strip() if not pd.isna(row['Customer/Project: Company Name']) else 'Unknown'
                    
                    try:
                        sales_price = float(row['Sales Price'])
                    except (ValueError, TypeError):
                        continue
                    
                    attribute_cost = 0
                    if process != 'LaserCut':
                        if process in pricing_rules["Process"] and step_process in pricing_rules["Process"][process]:
                            attribute_cost += pricing_rules["Process"][process][step_process]
                        if coating in pricing_rules["Coating"]:
                            attribute_cost += pricing_rules["Coating"][coating]
                    
                    base_cost = sales_price - attribute_cost
                    
                    results.append({
                        'Customer': customer,
                        'Frame': str(row['Frame']).strip(),
                        'Sales_Price': sales_price,
                        'Process': process,
                        'Step_Process': step_process,
                        'Coating': coating,
                        'Foil_Material': foil_material,
                        'Foil_Thickness': foil_thickness,
                        'Colour': colour,
                        'Attribute_Cost': attribute_cost,
                        'Base_Cost': base_cost
                    })
                except Exception as e:
                    continue
        
        if not results:
            logger.error("No valid data processed from Excel file")
            try:
                os.remove(file_path)
                logger.info(f"Removed invalid file: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
            return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">No valid data found in Excel file {os.path.basename(file_path)}. Please check the file contents.</p>')
        
        try:
            result_df = pd.DataFrame(results)
            logger.info(f"Processed {len(result_df)} rows before duplicate removal")
            if result_df.empty:
                logger.error("DataFrame is empty after processing")
                try:
                    os.remove(file_path)
                    logger.info(f"Removed invalid file: {file_path}")
                except Exception as e:
                    logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
                return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">No valid data after processing {os.path.basename(file_path)}. Please check the file contents.</p>')
            result_df['Customer'] = result_df['Customer'].astype(str)
            result_df['Base_Cost'] = pd.to_numeric(result_df['Base_Cost'], errors='coerce')
            if result_df['Base_Cost'].isna().all():
                logger.error("All Base_Cost values are invalid")
                try:
                    os.remove(file_path)
                    logger.info(f"Removed invalid file: {file_path}")
                except Exception as e:
                    logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
                return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">All Base_Cost values are invalid in {os.path.basename(file_path)}. Please check Sales Price data.</p>')
            material_columns = ['Customer', 'Process', 'Step_Process', 'Coating', 'Foil_Material', 'Foil_Thickness', 'Colour']
            result_df = result_df.loc[result_df.groupby(material_columns)['Base_Cost'].idxmin()].reset_index(drop=True)
            logger.info(f"After duplicate removal: {len(result_df)} unique combinations")
        except Exception as e:
            logger.error(f"Error processing results: {str(e)}")
            try:
                os.remove(file_path)
                logger.info(f"Removed invalid file: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
            return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Error processing results from {os.path.basename(file_path)}: {str(e)}. Please try again.</p>')
        
        try:
            if result_df.empty:
                chart_html = '<p class="error">No data available for chart. Please check your input data.</p>'
            elif result_df['Base_Cost'].isna().all():
                chart_html = '<p class="error">Invalid Base Cost data for chart. Please verify Sales Price and pricing rules.</p>'
            else:
                fig = px.bar(result_df, x='Customer', y='Base_Cost', title='Lowest Base Cost by Customer',
                             labels={'Base_Cost': 'Base Cost ($)', 'Customer': 'Customer'})
                fig.update_layout(xaxis_tickangle=45)
                chart_html = pio.to_html(fig, full_html=False)
                logger.info("Bar chart generated successfully")
        except Exception as e:
            logger.error(f"Error generating chart: {str(e)}")
            chart_html = f'<p class="error">Error generating chart: {str(e)}</p>'
        
        csv_path = os.path.join(UPLOAD_FOLDER, 'results.csv')
        excel_path = os.path.join(UPLOAD_FOLDER, 'results.xlsx')
        try:
            result_df.to_csv(csv_path, index=False)
            result_df.to_excel(excel_path, index=False, engine='openpyxl')
            logger.info(f"Results saved to {csv_path} and {excel_path}")
        except Exception as e:
            logger.error(f"Error saving results: {str(e)}")
            try:
                os.remove(file_path)
                logger.info(f"Removed invalid file: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
            return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Error saving results: {str(e)}. Please try again.</p>')
        
        try:
            os.remove(file_path)
            logger.info(f"Removed uploaded Excel file: {file_path}")
        except Exception as e:
            logger.warning(f"Failed to remove Excel file {file_path}: {str(e)}")
        
        logger.info("Rendering results page")
        return app.jinja_env.from_string(results_html).render(data=result_df.to_dict('records'), chart=chart_html)
    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}")
        try:
            os.remove(file_path)
            logger.info(f"Removed invalid file: {file_path}")
        except Exception as e:
            logger.warning(f"Failed to remove invalid file {file_path}: {str(e)}")
        return app.jinja_env.from_string(upload_html).render(error=f'<p class="error">Error reading Excel file {os.path.basename(file_path)}: {str(e)}. Please upload again.</p>')

@app.route('/download')
def download_csv():
    result_path = os.path.join(UPLOAD_FOLDER, 'results.csv')
    if os.path.exists(result_path):
        with open(result_path, 'rb') as f:
            response = make_response(f.read())
        response.headers['Content-Disposition'] = 'attachment; filename=results.csv'
        response.mimetype = 'text/csv'
        try:
            os.remove(result_path)
            logger.info(f"Deleted CSV after download: {result_path}")
        except Exception as e:
            logger.warning(f"Failed to delete CSV {result_path}: {str(e)}")
        return response
    logger.error("CSV file not found for download")
    return app.jinja_env.from_string(upload_html).render(error='<p class="error">No results available for download. Please process the file again.</p>')

@app.route('/download_excel')
def download_excel():
    result_path = os.path.join(UPLOAD_FOLDER, 'results.xlsx')
    if os.path.exists(result_path):
        with open(result_path, 'rb') as f:
            response = make_response(f.read())
        response.headers['Content-Disposition'] = 'attachment; filename=results.xlsx'
        response.mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        try:
            os.remove(result_path)
            logger.info(f"Deleted Excel after download: {result_path}")
        except Exception as e:
            logger.warning(f"Failed to delete Excel {result_path}: {str(e)}")
        return response
    logger.error("Excel file not found for download")
    return app.jinja_env.from_string(upload_html).render(error='<p class="error">No results available for download. Please process the file again.</p>')

if __name__ == '__main__':
    app.run(debug=True)