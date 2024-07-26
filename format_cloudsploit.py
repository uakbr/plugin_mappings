#!/usr/bin/env python3

import argparse, csv, json, pathlib, sys, xlsxwriter, zipfile

# Prepare global constants / variables
CHART_DATA = [{},{}] # Observation Areas (index 0), Risk Levels (index 1)
SLASH = '\\' if sys.platform == 'win32' else '/'
SUPPORTED_COMPLIANCE_STANDARDS = ['ALL', 'CMMC', 'CCPA', 'CIS Benchmarks', 'FedRamp', 'GDPR', 'HIPPA', 'ISO 27001', 'ISO 27017', 'ISO 27018', 'NIST 800-53', 'NIST 800-171', 'NIST CSF', 'PCI', 'SOC 2 Type II', 'SOC 3', 'Well Architected Framework']
SCAN_TYPE = 'cli'
MAPPINGS = {}
with open('./static/plugin_mappings.json') as file:
    MAPPINGS = json.load(file)


def append_row(domains, titles, risks, descriptions, efforts, assets, test_title, test_assets, compliance_standards):
    # Lookup CloudSploit plugin information from test title; Remove [DEFAULT] flag if it is present
    plugin_mappings = MAPPINGS.get(test_title, {})
    domain = plugin_mappings.get('PluginDomain', 'Unknown').replace('[DEFAULT]', 'Unknown')
    severity = plugin_mappings.get('PluginSeverity', 'Unknown').replace('[DEFAULT] ', '').capitalize()
    description = plugin_mappings.get('PluginTestDescription', 'Unknown').replace('[DEFAULT] ', '')
    remediation = plugin_mappings.get('PluginRecommendation', 'Unknown').replace('[DEFAULT] ', '')

    # Append values to the corresponding lists; empty current test assets afterwards
    domains.append(domain)
    titles.append(test_title)
    risks.append(severity.capitalize())
    descriptions.append(description)
    efforts.append(remediation)  # efforts.append(effort)
    assets.append('\n'.join(test_assets))
    
    # Update compliance standards (Columns H+)
    for standard in compliance_standards.keys():
        mappings = plugin_mappings.get('PluginComplianceMappings', {}).get(standard, [])
        compliance_standards[standard].append('\n'.join(mappings))

    # Update chart data
    CHART_DATA[0][domain] = CHART_DATA[0].get(domain, 0) + 1
    CHART_DATA[1][severity] = CHART_DATA[1].get(severity, 0) + 1

    # Return current number of rows in the worksheet
    return len(titles)


def format_sheet(worksheet, rows, formats, end_column, is_statistics=False):
    # Add Risk Level Colors
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Critical"', 'format': formats['Critical']})
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"High"', 'format': formats['High']})
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Moderate"', 'format': formats['Moderate']})  # Moderate check
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Medium"', 'format': formats['Moderate']})    # Medium check
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Low"', 'format': formats['Low']})
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Info"', 'format': formats['Info']})
    # Add Table Border Colors
    worksheet.conditional_format(f'A1:{end_column}1', {'type': 'formula', 'criteria': '=True', 'format': formats['Border']})
    worksheet.conditional_format(f'A1:A{rows+1}', {'type': 'formula', 'criteria': '=True', 'format': formats['Border']})
    worksheet.conditional_format(f'A{rows+1}:{end_column}{rows+1}', {'type': 'formula', 'criteria': '=True', 'format': formats['Border']})
    worksheet.conditional_format(f'{end_column}1:{end_column}{rows+1}', {'type': 'formula', 'criteria': '=True','format': formats['Border']})

    # Add Unknown Observation Domain / Risk Level Colors
    worksheet.conditional_format(f'B1:{end_column}{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Unknown"', 'format': formats['Unknown']})
    worksheet.conditional_format(f'D1:D{rows+1}', {'type': 'cell', 'criteria': '==', 'value': '"Unrated"', 'format': formats['Unknown']})
    # Alternate table row line color & set interior borders
    worksheet.conditional_format(f'B1:{end_column}{rows+1}', {'type': 'formula', 'criteria': '=ISODD(ROW())', 'format': formats['Row']})
    worksheet.conditional_format(f'B1:{end_column}{rows+1}', {'type': 'formula', 'criteria': '=True', 'format': formats['CellBorders']})
    # Set column widths
    worksheet.set_column(0, 0, 1)   # Left side table border
    worksheet.set_column(1, 2, 41)  # Report Observation Areas & Observation Titles
    worksheet.set_column(3, 3, 10)  # Risk Level
    if (is_statistics):
        worksheet.set_column(4, 6, 10)
    else:
        worksheet.set_column(4, 5, 40)  # Observation Description & Remediation Effort
        worksheet.set_column(6, 6, 100) # Affected Assets
    
    # Work around to compute how many compliance columns are needed based off the column letter
    number_compliance_columns = int.from_bytes(end_column.encode('utf-8'), 'little') - int.from_bytes('H'.encode('utf-8'), 'little')
    worksheet.set_column(7, 7+number_compliance_columns, 20)
    worksheet.set_column(f'{end_column}:{end_column}', 1)   # Right side border

    # Freeze top row (table headers) and set zoom to 75%
    worksheet.freeze_panes(1, 0)
    worksheet.set_zoom(90)
    return True


def draw_charts(workbook, chartsheet1, chartsheet2):
    # Enter chart data into the worksheet
    category_count = len(list(CHART_DATA[0].keys()))
    risk_type_count = len(list(CHART_DATA[1].keys()))
    
    # Create, fill out, and hide a worksheet to hold the chart information
    worksheet = workbook.add_worksheet('ChartData')
    worksheet.write_row('A1', ['Observation Category', 'Category Count', 'Observation Risk', 'Risk Count'])
    worksheet.write_column('A2', list(CHART_DATA[0].keys()))
    worksheet.write_column('B2', list(CHART_DATA[0].values()))
    worksheet.write_column('C2', list(CHART_DATA[1].keys()))
    worksheet.write_column('D2', list(CHART_DATA[1].values()))
    #worksheet.hide()

    # Draw Observation Domains Chart
    by_category = workbook.add_chart({'type': 'pie'})
    by_category.add_series({
        'name':       'Observation Domain Count',
        'categories': ['ChartData', 1, 0, category_count, 0],
        'values':     ['ChartData', 1, 1, category_count, 1]
    })
    by_category.set_title({'name': 'Unique Observations by Domain'})
    by_category.set_style(2)

    # Format Risk Levels chart
    by_risk = workbook.add_chart({'type': 'bar'})
    by_risk.add_series({
        'name': 'Observation Risk Levels',
        'categories': ['ChartData', 1, 2, risk_type_count, 2],
        'values':     ['ChartData', 1, 3, risk_type_count, 3]
    })
    by_risk.set_title({'name': 'Unique Observations by Risk Level'})
    by_risk.set_legend({'none': True})
    by_risk.set_style(2)

    # Assign charts to their corresponding sheet
    chartsheet1.set_chart(by_category)
    print('[+] Charted unique observations by test domain')
    chartsheet2.set_chart(by_risk)
    print('[+] Charted unqiue observations by test severity level')
    return True


def format_observations(workbook, filename, sheetname, standards, formats, accumulator = 0):
    # Ensure that the target file is indeed a CSV
    if filename[1+filename.rfind('.')::].lower() != 'csv':
        print('[!] File not CSV, skipping!')
        return False
    
    # These variables represent which column the data is found in within cloudsploit CSVs ()
    title, asset, region, result = (1, 3, 4, 5) if SCAN_TYPE == 'cli' else (1, 5, 3, 4)
    print(f'[+] Writing observations sheet to \'{sheetname}\' tab')
    try:
        # Add worksheet to the xlsx file; Initialize number of rows in the worksheet
        worksheet = workbook.add_worksheet(sheetname)
        row_count = 0
        with open(filename) as file:
            lines = csv.reader(file)
            
            # Worksheet columns are abstracted as lists Column A/H are table borders
            domains =   ['Report Observation Domain']           # Column B, Observation domain assigned to a test
            titles =  ['Observation Title']                     # Column C, Title of the test
            risks = ['Risk Level']                              # Column D, Default risk level for the test before consultant review
            descriptions = ['Report Observation Description']   # Column E, Observation descriptions (To be filled in by consultant)
            efforts = ['Remediation Effort']                    # Column F, Remediation efforts (to be filled in by consultant)
            assets = []                                         # Column G, Assets affected (or region if no asset specified) by the failed test
            
            compliance_standards = {}                           # Columns after G - compliance mappings
            for standard in standards:
                compliance_standards[standard] = [standard]     # List with first element as the compliance standard's name

            current_assets = []     # Temp container for aggregating assets affected by a single test
            prior_test = ''         # Temp variable to hold previous test title
            initialized = False     # Used to make sure we fetch the first Test title and don't append an empty row
            lines.__next__()        # Skip first row (CSV headers)
            for entry in lines:
                if entry[result] == 'FAIL':
                    # Make sure we get the first failed test title and don't append an empty row
                    if not initialized:
                        prior_test = entry[title]
                        initialized = True

                    # Check if current entry is for a new test
                    if entry[title] != prior_test:
                        append_row(domains, titles, risks, descriptions, efforts, assets, prior_test, current_assets, compliance_standards)

                        current_assets.clear()
                        prior_test = entry[title]
                        
                    # Get the asset (if applicable) or region affected (if no asset listed)
                    current_assets.append(entry[region]) if (entry[asset] == 'N/A') else current_assets.append(entry[asset])
            # Must write final row after the loop breaks
            row_count = append_row(domains, titles, risks, descriptions, efforts, assets, prior_test, current_assets, compliance_standards)
            
            # Write parsed data to the worksheet
            worksheet.write_column('B1', domains, formats['Text'])
            worksheet.write_column('C1', titles, formats['Text'])
            worksheet.write_column('D1', risks, formats['Center'])
            worksheet.write_column('E1', descriptions, formats['Text'])
            worksheet.write_column('F1', efforts, formats['Text'])
            worksheet.write_column('G1', ['Affected Assets'])
            worksheet.write_column('G2', assets, formats['Assets']) # End with()
            
            # Now generate the compliance standard columns TODO: Figure out how to write compiance mappings by row
            next_column = int.from_bytes('H'.encode('utf-8'), 'little')
            for standard in compliance_standards.keys():
                worksheet.write_column(f'{chr(next_column)}1', compliance_standards[standard], formats['Assets'])
                next_column += 1
            table_end_column = chr(next_column)

        return format_sheet(worksheet, row_count, formats, table_end_column)

    except xlsxwriter.exceptions.InvalidWorksheetName: # Probably too long (>31 characters)
        print(f'[!] Sheet \'{sheetname}\' has  invalid name! Attempting to trim')
        if len(sheetname) > 31:
            sheetname = sheetname[0:31]
        else:
            sheetname = input(f'[!] Sheet \'{sheetname}\' is still invalid! Enter new valid name for the sheet: ')
        return format_observations(workbook, filename, sheetname, standards, formats)
    except xlsxwriter.exceptions.DuplicateWorksheetName: # Sheet with the same name already created
        print(f'[!] Duplicate sheet found for \'{sheetname}\'! Attempting to mark duplicate!')
        accumulator += 1
        sheetname = sheetname[0:len(sheetname)-(len(str(accumulator))+2)] # Must account for multi-digit numbers as well as the two parenthesis
        sheetname += f'({accumulator})'
        return format_observations(workbook, filename, sheetname, standards, formats, accumulator)
    except xlsxwriter.exceptions.FileCreateError:
        print(f'[x] Could not create file {filename}!')
        exit()


def format_statistics(workbook, filename, sheetname, standards, formats):
    # Ensure that the target file is indeed a CSV
    if filename[1+filename.rfind('.')::].lower() != 'csv':
        print('[!] File not CSV, skipping!')
        return False

    print(f'[+] Computing pass/fail rates in \'Pass-Fail Rates\' tab')
    # These variables represent which column the data is found in within cloudsploit CSVs ()
    title, asset, region, result = (1, 3, 4, 5) if SCAN_TYPE == 'cli' else (1, 5, 3, 4)
    passing = 'OK' if SCAN_TYPE == 'cli' else 'PASS'
    try:
        worksheet = workbook.add_worksheet(f'Pass-Fail Rates')
        lines = None
        domains = ['Report Observation Domain'] # Column B, category of test; same as in observtaion sheet
        titles = ['Observation Title']          # Column C, title of test; same as in observation sheet
        risks = ['Risk Level']                  # Column D, severity of failed test; same as in observation sheet
        pass_count = ['Success Count']          # Column E, number of instances where the test returned a passing result
        fail_count = ['Fail Count']             # Column F, number of instances where test returned a failed result
        success_rate = ['Pass Rate']            # Column G, percentage of pass / fail rate
            
        compliance_standards = {}                           # Columns after G - compliance mappings
        for standard in standards:
            compliance_standards[standard] = [standard]     # List with first element as the compliance standard's name

        with open(filename) as file:
            lines = csv.reader(file)

            prior_test = ''         # Temp variable to hold previous test title
            initialized = False     # Used to make sure we fetch the first Test title and don't append an empty row
            current_passes, current_fails, total_entries_for_test, row_count = 0, 0, 0, 0
            lines.__next__()        # Skip first row (CSV headers)
            for entry in lines:
                # Make sure we get the first failed test title and don't append an empty row
                if not initialized:
                    prior_test = entry[title]
                    initialized = True
                
                if entry[title] != prior_test: # Went through all of the current test results; add to worksheet & reset counters
                    plugin_mappings = MAPPINGS.get(prior_test, {})
                    domain = plugin_mappings.get('PluginDomain', 'Unknown').replace('[DEFAULT]', 'Unknown')
                    severity = plugin_mappings.get('PluginSeverity', 'Unknown').replace('[DEFAULT] ', '').capitalize()

                    # Add worksheet row entries
                    domains.append(domain)
                    titles.append(prior_test)
                    risks.append(severity)
                    pass_count.append(f'{current_passes}')
                    fail_count.append(f'{current_fails}')
                    rate = (current_passes / total_entries_for_test)
                    success_rate.append(rate)
                    row_count += 1 # Rows on final statistics sheet

                    # Add compliance mappings
                    for standard in compliance_standards.keys():
                        mappings = plugin_mappings.get('PluginComplianceMappings', {}).get(standard, [])
                        compliance_standards[standard].append('\n'.join(mappings))

                    # Reset counters & update prior test name
                    current_passes, current_fails, total_entries_for_test = 0, 0, 0
                    prior_test = entry[title]

                # Count pass / fail / total entries for each test
                total_entries_for_test += 1
                if entry[result] == 'FAIL':
                    current_fails += 1
                if entry[result] == passing:
                    current_passes += 1
            # END LOOP
            
            # After loop breaks, must account for the last test run - I know, horrible coppy / paste practice especially since its so similar to append_row() :(
            plugin_mappings = MAPPINGS.get(prior_test, {})
            domain = plugin_mappings.get('PluginDomain', 'Unknown').replace('[DEFAULT]', 'Unknown')
            severity = plugin_mappings.get('PluginSeverity', 'Unknown').replace('[DEFAULT] ', '').capitalize()
            domains.append(domain)
            titles.append(entry[title])
            risks.append(severity)
            pass_count.append(f'{current_passes}')
            fail_count.append(f'{current_fails}')
            rate = (current_passes / total_entries_for_test)
            success_rate.append(rate)
                    
            # Add compliance mappings for the final test
            for standard in compliance_standards.keys():
                mappings = plugin_mappings.get('PluginComplianceMappings', {}).get(standard, [])
                compliance_standards[standard].append('\n'.join(mappings))
            row_count += 2 # Account for final row + the border row

            # Now actually write the data stored within our lists to the worksheet
            worksheet.write_column('B1', domains, formats['Text'])
            worksheet.write_column('C1', titles, formats['Text'])
            worksheet.write_column('D1', risks, formats['Center'])
            worksheet.write_column('E1', pass_count, formats['Text'])
            worksheet.write_column('F1', fail_count, formats['Text'])
            worksheet.write_column('G1', success_rate, formats['Text'])
                
            # Now generate the compliance standard columns TODO: Figure out how to write compiance mappings by row
            next_column = int.from_bytes('H'.encode('utf-8'), 'little')
            for standard in compliance_standards.keys():
                worksheet.write_column(f'{chr(next_column)}1', compliance_standards[standard], formats['Assets'])
                next_column += 1
            table_end_column = chr(next_column)
            
            return format_sheet(worksheet, row_count, formats, table_end_column, True) # True because writing statistics sheet

    except xlsxwriter.exceptions.InvalidWorksheetName or xlsxwriter.exceptions.DuplicateWorksheetName as err:
        print(f'[x] {err}')
    return True


def get_targets(target_file):
    targets = []
    with open(target_file) as file:
        for target in file:
            targets.append(target)
    return targets


def get_targets_recursive(directory):
    paths = list(pathlib.Path(directory).rglob('*.[cC][sS][vV]'))
    targets = []
    for target in paths:
        targets.append(str(target))    # Cast WindowsPath to string
    return targets


def print_usage():
    print('Please enter either a single \'-l\' OR \'-t\' OR \'-d\' option. Optionally, specify the output filename with the \'-o\' flag.')
    print('\tExample: py format_cloudsploit.py -l list_of_targets.txt')
    print('\tExample: py format_cloudsploit.py -t C:\\Users\\Me\\Desktop\\cloudsploit_output.csv -o CloudsploitObservations.xlsx')
    print('\tExample: py format_cloudsploit.py -d . -o AllScans.xlsx')
    print('Optionally, add the \'-z\' flag to compress the resulting output into an additional zipped file.')
    print('\tExample: py format_cloudsploit.py -d C:\\Directory\\With\\Lots\\Of\\CSVs -z -o LargeFile.xlsx\n')    
    exit()


def compress_file(filename):
    print(f'[!] Attempting to compress file \'{filename}\'')
    compression = zipfile.ZIP_DEFLATED
    zipped = zipfile.ZipFile(filename+'.zip', mode='w')
    try:
        zipped.write(filename, filename, compress_type=compression)
    except FileNotFoundError:
        print(f'[x] Error compressing file {filename}! Exiting!')
    finally:
        print(f'[=] Finished file compression! Wrote to {filename+".zip"}!')
        zipped.close()
    exit()


def add_formats(workbook):
    gray = workbook.add_format({'bg_color':'#C0C0C0'})         # Light Grey (Overall Table Rows)    
    grey = workbook.add_format({
        'bg_color':'#404040',
        'bold':1,
        'font_color':'white'   
    })                                                          # Dark Grey (Overall Table Borders)
    salmon = workbook.add_format({'bg_color': '#FF0066'})       # Salmon (Unknown Observation Areas)     
    dark_red = workbook.add_format({'bg_color': '#880000'})     # Dark Red (Critical Risk)  
    orange = workbook.add_format({'bg_color': '#AA5500'})       # Orange (High Risk)   
    yellow = workbook.add_format({'bg_color': '#FFFF66'})       # Yellow (Moderate Risk)   
    green = workbook.add_format({'bg_color': '#009900'})        # Green (Low Risk)    
    blue = workbook.add_format({'bg_color': '#0099AA'})         # Blue (Informational)
    assets = workbook.add_format({'font_size':9})               # Small Text (Affected Assets)
    assets.set_text_wrap()
    assets.set_align('left')
    assets.set_align('top')
    text = workbook.add_format()                                # Regular Text (Interior Rows)
    text.set_text_wrap()
    text.set_align('left')
    text.set_align('top')
    cell_borders = workbook.add_format({'border':1})            # Add Cell Borders
    center = workbook.add_format()                                # Centered Text
    center.set_align('center')
    center.set_align('vcenter')
    return {
        'Row': gray,
        'Border': grey,
        'Unknown' : salmon,
        'Critical': dark_red,
        'High': orange,
        'Moderate': yellow,
        'Low': green,
        'Info': blue,
        'Assets': assets,
        'Text': text,
        'CellBorders': cell_borders,
        'Center': center
    }


def copy_raw_output(workbook, filename):
    worksheet = workbook.add_worksheet('Raw Output')
    with open(filename) as file:
        lines = csv.reader(file)
        row = 1
        for entry in lines:
            worksheet.write_row(f'A{row}', entry)
            row += 1
    return True


def format_cloudsploit(args):
    workbook = xlsxwriter.Workbook(args.output.strip())
    observation_categories_chart = workbook.add_chartsheet('Observation Categories')
    risk_levels_chart = workbook.add_chartsheet('Risk Levels')
    formats = add_formats(workbook)
    
    # Determine scans to include
    targets = []
    if (args.target): # individual CSV
        targets.append(args.target)
    elif (args.directory):  # Directory specified - get all CSVs recursively
        targets = get_targets_recursive(args.directory)
    elif (args.list): # List of CSVs specified in a file
        targets = get_targets(args.list)
    else:
        print_usage()
    
    # Determine compliance mappings
    standards = []
    if args.compliance:
        if (args.compliance == 'ALL'):
            standards = SUPPORTED_COMPLIANCE_STANDARDS
            standards.remove('ALL')
        else:
            standards = args.compliance.split(',')
            for standard in standards:
                standard = standard.strip()
                if standard not in SUPPORTED_COMPLIANCE_STANDARDS:
                    print(f'[x] Supported standards are: {", ".join(SUPPORTED_COMPLIANCE_STANDARDS)}')
                    print(f'[x] Compliance standard \'{standard}\' not supported! Exiting!')
                    exit()

    worksheet_count = 0
    for target in targets:
        sheet = target[1+target.rfind(SLASH):target.rfind('.')].capitalize()          # Strip directory path & extension
        print(f'[+] Creating cloud scan workbook for {target.strip()}...')
        format_observations(workbook, target.strip(), sheet, standards, formats)       # Convert to Excel format
        worksheet_count += 1
        if (args.include_statistics):
            format_statistics(workbook, target.strip(), sheet, standards, formats)
            worksheet_count += 1
    
    # Now Compute Charts
    print(f'[+] Computing charts in \'{args.output}\' for resulting observations...')
    draw_charts(workbook, observation_categories_chart, risk_levels_chart)
    print(f'[+] Finished charting observation data!')
    worksheet_count += 3

    if (args.target):
        print(f'[+] Copying raw Cloudsploit results to \'Raw Output\' tab')
        copy_raw_output(workbook, args.target)
        worksheet_count += 1

    workbook.close()
    print(f'[!] Note: You\'ll still need to account for \'Unknown\' values, sort observations and statistics tabs and update the ChartData and RiskLevels tabs to have proper numbers / coloring!')
    print(f'[=] Finished parsing CSVs! Wrote {worksheet_count} worksheets to new workbook, {args.output}')
    compress_file(args.output) if (args.zip) else exit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser('python3 format_cloudsploit.py -d C:\Path\To\Directory -o PreliminaryObservations.xlsx')
    parser.add_argument('-l', '--list', help='File containing a list of input CSVs, one per line', required=False)
    parser.add_argument('-t', '--target', help='Target input CSV file', required=False)
    parser.add_argument('-d', '--directory', help='Target folder. FormatCloudsploit will recursively search for all CSV files and merge them into one Excel workbook', required=False)
    parser.add_argument('-z', '--zip', help='Create a compressed zip file as well the original (default: False)', default=False, action='store_true', required=False)
    parser.add_argument('-o', '--output', help='Filename to write to (default: \'observations.xlsx\')', default='observations.xlsx', required=False)
    parser.add_argument('-c', '--compliance', help='Compliance standard to map results (default: ALL)', default="ALL", required=False)
    parser.add_argument('-a', '--aquawave', action='store_true', default=False, help='Indicates CSV results from Aquawave rather than CLI (default: False)')
    parser.add_argument('--include-statistics', action='store_true', default=False, help='Include pass-fail rates in a seperate tab (default: False)')
    args = parser.parse_args()

    if (args.aquawave):
        SCAN_TYPE = 'aquawave'

    if (not any([args.target, args.list, args.directory])):
        print_usage()
    else:
        format_cloudsploit(args)

