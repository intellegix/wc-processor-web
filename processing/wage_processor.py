"""
Workers Compensation Wage Processor
Core data processing engine for wage report validation and transformation
"""

import pandas as pd
import os
import re
from datetime import datetime


def exclude_cy_jobs(df):
    """Exclude any rows where job_no starts with 'CY' and two digits"""
    cy_job_pattern = re.compile(r'^CY\d{2}')
    mask = ~df['job_no'].astype(str).str.strip().str.upper().str.match(cy_job_pattern)
    return df[mask]


def get_4digit_to_6digit_mapping():
    """
    Convert standard 4-digit NCCI codes to California 6-digit codes.
    """
    return {
        5403: 540321,   # Carpentry Low Wage (< $41/hr)
        5432: 543221,   # Carpentry High Wage (>= $41/hr)
        5446: 544615,   # Wallboard Low Wage (< $41/hr)
        5447: 544715,   # Wallboard High Wage (>= $41/hr)
        5482: 548234,   # Painting High Wage (>= $32/hr)
        5485: 548515,   # Plastering High Wage (>= $38/hr)
        5553: 555311,   # Roofing High Wage (>= $31/hr)
        8810: 881002,   # Clerical Office Employees (convert to 6-digit)
        # Additional codes
        510704: 510704, # Door/Window Installation (already 6-digit)
        553823: 553823, # Sheet Metal Low Wage (already 6-digit)
        554222: 554222, # Sheet Metal High Wage (already 6-digit)
        621837: 621837, # Excavation Low Wage (already 6-digit)
        622038: 622038, # Excavation High Wage (already 6-digit)
        822704: 822704, # Construction Yards/Shops (already 6-digit)
        874210: 874210, # Salespersons Outside (already 6-digit)
        881002: 881002, # Clerical (already 6-digit)
    }


def get_class_code_mapping():
    """
    Define the high-wage to low-wage class code mappings for California Workers' Comp.
    """
    return {
        543221: 540321,  # Carpentry High Wage -> Carpentry Low Wage
        544715: 544615,  # Wallboard High Wage -> Wallboard Low Wage
        548234: 547434,  # Painting High Wage -> Painting Low Wage
        548515: 548415,  # Plastering High Wage -> Plastering Low Wage
        554222: 553823,  # Sheet Metal High Wage -> Sheet Metal Low Wage
        555311: 555211,  # Roofing High Wage -> Roofing Low Wage
        622038: 621837,  # Excavation High Wage -> Excavation Low Wage
    }


def get_wage_thresholds():
    """
    Define wage thresholds for dual-wage classifications.
    """
    return {
        'carpentry': {
            'high_code': 543221,
            'low_code': 540321,
            'threshold': 41.00,
            'name': 'Carpentry'
        },
        'wallboard': {
            'high_code': 544715,
            'low_code': 544615,
            'threshold': 41.00,
            'name': 'Wallboard'
        },
        'painting': {
            'high_code': 548234,
            'low_code': 547434,
            'threshold': 32.00,
            'name': 'Painting'
        },
        'plastering': {
            'high_code': 548515,
            'low_code': 548415,
            'threshold': 38.00,
            'name': 'Plastering/Stucco'
        },
        'sheet_metal': {
            'high_code': 554222,
            'low_code': 553823,
            'threshold': 33.00,
            'name': 'Sheet Metal'
        },
        'roofing': {
            'high_code': 555311,
            'low_code': 555211,
            'threshold': 31.00,
            'name': 'Roofing'
        },
        'excavation': {
            'high_code': 622038,
            'low_code': 621837,
            'threshold': 40.00,
            'name': 'Excavation'
        }
    }


def convert_4digit_to_6digit_codes(df):
    """
    Convert 4-digit NCCI codes to 6-digit California codes.
    """
    code_map = get_4digit_to_6digit_mapping()
    conversions = []

    for idx in df.index:
        try:
            original_code = float(df.at[idx, 'Cost Code'])

            if original_code in code_map:
                new_code = code_map[original_code]

                if original_code != new_code:
                    df.at[idx, 'Cost Code'] = new_code
                    conversions.append({
                        'original': int(original_code),
                        'new': new_code,
                        'employee': df.at[idx, 'Employee Name'],
                        'earnings': df.at[idx, 'Earnings']
                    })
        except (ValueError, KeyError, TypeError):
            continue

    return df, conversions


def apply_employee_specific_corrections(df):
    """
    Apply specific employee class code corrections.
    """
    corrections = []

    employee_rules = {
        'Kidwell , Austin': {
            'correct_code': 881002,
            'reason': 'Office/clerical duties only',
            'description': 'Clerical Office Employees (NOC)'
        }
    }

    for employee_name, rule in employee_rules.items():
        employee_mask = df['Employee Name'].str.strip() == employee_name.strip()

        if employee_mask.any():
            original_codes = df.loc[employee_mask, 'Cost Code'].unique()

            for original_code in original_codes:
                try:
                    if float(original_code) != rule['correct_code']:
                        code_mask = employee_mask & (df['Cost Code'] == original_code)
                        count = code_mask.sum()
                        total_earnings = df.loc[code_mask, 'Earnings'].sum()

                        df.loc[code_mask, 'Cost Code'] = rule['correct_code']

                        corrections.append({
                            'employee': employee_name,
                            'original_code': int(float(original_code)),
                            'corrected_code': rule['correct_code'],
                            'reason': rule['reason'],
                            'count': count,
                            'earnings': total_earnings
                        })
                except (ValueError, TypeError):
                    continue

    return df, corrections


def validate_and_correct_all_class_codes(df):
    """
    COMPREHENSIVE CLASS CODE VALIDATION - Checks EVERY transaction.
    """
    thresholds = get_wage_thresholds()
    class_code_map = get_class_code_mapping()

    code_to_trade = {}

    for trade_key, trade_info in thresholds.items():
        high_code = trade_info['high_code']
        low_code = trade_info['low_code']
        threshold = trade_info['threshold']
        name = trade_info['name']

        if high_code:
            code_to_trade[high_code] = {
                'type': 'high',
                'threshold': threshold,
                'name': name,
                'low_code': low_code,
                'high_code': high_code
            }
        if low_code:
            code_to_trade[low_code] = {
                'type': 'low',
                'threshold': threshold,
                'name': name,
                'low_code': low_code,
                'high_code': high_code
            }

    corrections = []
    validation_summary = {
        'total_transactions': len(df),
        'validated': 0,
        'corrected': 0,
        'drive_time_corrected': 0,
        'wage_corrected': 0,
        'skipped': 0
    }

    regular_wage_types = ['REG', 'VAC', 'SICK', 'DBA', 'SUPP', 'SAL', 'OSAL', 'PWREG']

    for idx in df.index:
        try:
            earnings = float(df.at[idx, 'Earnings'])
            hours = float(df.at[idx, 'Hours'])
            cost_code = float(df.at[idx, 'Cost Code'])
            earn_type = df.at[idx, 'Earn Type'].upper().strip()
            employee = df.at[idx, 'Employee Name']
            job_no = df.at[idx, 'Job No']

            if hours > 0:
                actual_rate = earnings / hours
            else:
                validation_summary['skipped'] += 1
                continue

            # VALIDATION RULE 1: DRIVE TIME - MUST use low wage code
            if earn_type in ['DRIVE', 'DROVT']:
                if cost_code in class_code_map:
                    correct_code = class_code_map[cost_code]
                    df.at[idx, 'Cost Code'] = correct_code
                    corrections.append({
                        'row': idx,
                        'employee': employee,
                        'job_no': job_no,
                        'earn_type': earn_type,
                        'hours': hours,
                        'rate': actual_rate,
                        'original_code': int(cost_code),
                        'corrected_code': correct_code,
                        'reason': 'DRIVE TIME must use low wage code',
                        'category': 'DRIVE_TIME',
                        'earnings': earnings
                    })
                    validation_summary['corrected'] += 1
                    validation_summary['drive_time_corrected'] += 1
                validation_summary['validated'] += 1
                continue

            # VALIDATION RULE 2: WAGE-BASED (regular wages only)
            if earn_type not in regular_wage_types:
                validation_summary['validated'] += 1
                continue

            if cost_code in code_to_trade:
                trade = code_to_trade[cost_code]
                threshold = trade['threshold']
                low_code = trade['low_code']
                high_code = trade['high_code']
                trade_name = trade['name']

                if actual_rate >= threshold:
                    if cost_code == low_code:
                        df.at[idx, 'Cost Code'] = high_code
                        corrections.append({
                            'row': idx,
                            'employee': employee,
                            'job_no': job_no,
                            'earn_type': earn_type,
                            'hours': hours,
                            'rate': actual_rate,
                            'original_code': int(cost_code),
                            'corrected_code': high_code,
                            'reason': f'{trade_name}: ${actual_rate:.2f}/hr >= ${threshold:.2f}/hr threshold (should be HIGH wage)',
                            'category': 'WAGE_BASED_UP',
                            'earnings': earnings
                        })
                        validation_summary['corrected'] += 1
                        validation_summary['wage_corrected'] += 1
                else:
                    if cost_code == high_code:
                        df.at[idx, 'Cost Code'] = low_code
                        corrections.append({
                            'row': idx,
                            'employee': employee,
                            'job_no': job_no,
                            'earn_type': earn_type,
                            'hours': hours,
                            'rate': actual_rate,
                            'original_code': int(cost_code),
                            'corrected_code': low_code,
                            'reason': f'{trade_name}: ${actual_rate:.2f}/hr < ${threshold:.2f}/hr threshold (should be LOW wage)',
                            'category': 'WAGE_BASED_DOWN',
                            'earnings': earnings
                        })
                        validation_summary['corrected'] += 1
                        validation_summary['wage_corrected'] += 1

            validation_summary['validated'] += 1

        except (ValueError, KeyError, TypeError):
            validation_summary['skipped'] += 1
            continue

    report = {
        'summary': validation_summary,
        'corrections': corrections
    }

    return df, report


def load_and_process_wage_report(file_path, output_folder, report_name="ASRWorkersCompReport.csv", include_subtotals=True):
    """Main function to load and process wage reports"""

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    file_ext = os.path.splitext(file_path)[1].lower()
    try:
        if file_ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, sheet_name=0)
    except Exception as e:
        raise RuntimeError(f"Error loading file: {e}")

    # Check if already processed
    processed_report_columns = {'Employee Name', 'Employee Number', 'Job No', 'Cost Code', 'Earn Type', 'Hours', 'Earnings'}
    if processed_report_columns.issubset(set(df.columns)):
        print(f"[INFO] File appears to be an already-processed workers comp report")
        return df, file_path

    # Check for raw payroll export
    core_required_columns = {'emp_name', 'employee_no', 'job_desc', 'class', 'earn_type_no', 'hours', 'earnings', 'job_no'}
    missing_core = core_required_columns - set(df.columns)
    if missing_core:
        raise ValueError(f"Invalid file format. Missing required columns: {missing_core}")

    # Add optional columns with default values
    if 'exposure' not in df.columns:
        df['exposure'] = df['earnings']

    if 'sort_option' not in df.columns:
        df['sort_option'] = df['job_desc']

    # Define wage types
    regular_wages = ['REG', 'VAC', 'BON', 'SUPP', 'SICK', 'DBA', 'DRIVE', 'OSAL', 'SAL', 'PWREG']
    overtime_wages = ['OVT', 'DROVT', 'PWOT']
    doubletime_wages = ['DBL']
    target_earn_types = regular_wages + overtime_wages + doubletime_wages

    # Clean and filter data
    df['earn_type_no_clean'] = df['earn_type_no'].astype(str).str.strip().str.upper()
    mask = df['earn_type_no_clean'].isin(target_earn_types)
    filtered = df.loc[mask].copy()

    # Exclude CY jobs
    filtered = exclude_cy_jobs(filtered)

    # Prepare columns for reporting
    columns = ['emp_name', 'employee_no', 'job_no', 'job_desc', 'class', 'earn_type_no_clean', 'hours', 'earnings', 'exposure', 'sort_option']
    detail = filtered.loc[:, columns].copy()
    detail = detail.rename(columns={
        'emp_name': 'Employee Name',
        'employee_no': 'Employee Number',
        'job_no': 'Job No',
        'job_desc': 'Job Description',
        'class': 'Cost Code',
        'earn_type_no_clean': 'Earn Type',
        'hours': 'Hours',
        'earnings': 'Earnings',
        'exposure': 'Exposure',
        'sort_option': 'Sort Option'
    })

    # Clean up datatypes
    for col in ['Cost Code', 'Sort Option']:
        detail[col] = detail[col].astype(str).str.strip()
    detail['Employee Number'] = detail['Employee Number'].astype(str).str.strip()
    for col in ['Hours', 'Earnings', 'Exposure']:
        detail[col] = pd.to_numeric(detail[col], errors='coerce').round(2)
    detail = detail.fillna({'Hours': 0, 'Earnings': 0, 'Exposure': 0})

    # Convert 4-digit to 6-digit codes
    detail, code_conversions = convert_4digit_to_6digit_codes(detail)

    # Apply employee-specific corrections
    detail, emp_corrections = apply_employee_specific_corrections(detail)

    # Comprehensive class code validation
    detail, validation_report = validate_and_correct_all_class_codes(detail)

    # Add detail rows with individual-level grand totals (conditional)
    detail = detail.sort_values(by=['Employee Number', 'Employee Name', 'Sort Option', 'Job No', 'Job Description', 'Earn Type'])

    if include_subtotals:
        all_rows = []
        for emp, emp_group in detail.groupby('Employee Name', sort=True):
            emp_number = emp_group['Employee Number'].iloc[0]

            for _, row in emp_group.iterrows():
                all_rows.append(row)

            # Subtotals
            subgroup_types = {
                'Regular Wages': regular_wages,
                'Overtime Wages': overtime_wages,
                'Doubletime Wages': doubletime_wages
            }
            for name, types in subgroup_types.items():
                sub = emp_group[emp_group['Earn Type'].isin(types)]
                if not sub.empty:
                    subtotal = {
                        'Employee Name': emp,
                        'Employee Number': emp_number,
                        'Job No': '',
                        'Job Description': f'---{name.upper()} TOTAL---',
                        'Cost Code': '',
                        'Earn Type': ','.join(types),
                        'Hours': sub['Hours'].sum().round(2),
                        'Earnings': sub['Earnings'].sum().round(2),
                        'Exposure': sub['Exposure'].sum().round(2),
                        'Sort Option': ''
                    }
                    all_rows.append(pd.Series(subtotal))

            # Grand total
            emp_total = {
                'Employee Name': emp,
                'Employee Number': emp_number,
                'Job No': '',
                'Job Description': '--GRAND TOTAL FOR EMPLOYEE--',
                'Cost Code': '',
                'Earn Type': '',
                'Hours': emp_group['Hours'].sum().round(2),
                'Earnings': emp_group['Earnings'].sum().round(2),
                'Exposure': emp_group['Exposure'].sum().round(2),
                'Sort Option': ''
            }
            all_rows.append(pd.Series(emp_total))

        report_df = pd.DataFrame(all_rows, columns=detail.columns)
    else:
        # Return only detail rows without subtotals for clean combining
        report_df = detail.copy()

    # Output file management
    os.makedirs(output_folder, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_folder, f"{timestamp}_{report_name}")
    report_df.to_csv(output_path, index=False)

    print(f"Report saved: {output_path}")
    print(f"Records processed: {len(detail)}")

    return report_df, output_path
