import csv
import datetime
import tempfile
import zipfile

import xlsxwriter
import pandas as pd
import operator

def ingest_files(file_paths):
    """
    Accepts a list of .csv and/or .zip files.
    Extracts CSVs if needed and returns combined parsed rows.
    """
    combined_rows = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for path in file_paths:
            if path.lower().endswith(".csv"):
                combined_rows.extend(csv_import(path))

            elif path.lower().endswith(".zip"):
                with zipfile.ZipFile(path, 'r') as z:
                    for name in z.namelist():
                        if name.lower().endswith(".csv"):
                            extracted = z.extract(name, tmpdir)
                            combined_rows.extend(csv_import(extracted))

    return combined_rows

def normalize_datetime(value):
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%y %H:%M"):
        try:
            return datetime.datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(f"Invalid date format: {value}")

def csv_import(filepath):
    expected_headers = [
        'location', 'sublocation', 'associate_vlan', 'device_mac', 'client_mac',
        'start_time', 'end_time', 'client_ip', 'client_host_name', 'client_os_name',
        'bssid', 'ssid'
    ]
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for row in reader:
            if set(expected_headers[:4]).issubset(row):
                headers = row
                break
        header_map = {key: headers.index(key) for key in expected_headers if key in headers}
        rows = []
        for row in reader:
            if len(row) < len(header_map):
                continue
            entry = {key: row[idx].strip() for key, idx in header_map.items()}
            rows.append(entry)
    return rows

def generate_excel_report(data, selected_sites, output_path, date_from=None, date_to=None, aggregate_floors=False, tab_per_building = False):

    df = pd.DataFrame(data)
    df['start_time'] = df['start_time'].apply(normalize_datetime)
    df['end_time'] = df['end_time'].apply(normalize_datetime)
    df['connected_time'] = (df['end_time'] - df['start_time']).dt.total_seconds()
    df['session_date'] = df['end_time']

    # building is the part before the pipe: "building|floor"
    # if no pipe exists, the whole sublocation is treated as building
    df["building"] = df["sublocation"].astype(str).str.split("|", n=1).str[0].str.strip()

    if date_from:
        df = df[df['session_date'] >= date_from]
    if date_to:
        df = df[df['session_date'] <= date_to]

    # Order of dates for consistent columns
    date_cols = sorted(df['session_date'].dt.normalize().unique())

    workbook = xlsxwriter.Workbook(output_path)

    # Aggregate report (multiple sites)
    if len(selected_sites) > 1:
        agg_df = df[df['location'].isin(selected_sites)]
        generate_site_report(
            agg_df,
            "Report",
            workbook,
            date_cols,
            aggregate_floors=aggregate_floors
        )

    # Per-site reports
    for site in selected_sites:
        site_df = df[df['location'] == site]
        if not site_df.empty:
            generate_site_report(
                site_df,
                site,
                workbook,
                date_cols,
                aggregate_floors=aggregate_floors
            )

    # Per-building reports (ONLY when aggregating floors)
    if aggregate_floors and tab_per_building:
        # Only buildings that appear in selected sites
        building_df = df[df['location'].isin(selected_sites)]

        for building in sorted(building_df['building'].dropna().unique()):
            bldg_df = building_df[building_df['building'] == building]

            if bldg_df.empty:
                continue

            # Sheet names must be <= 31 chars and unique
            sheet_name = f"Bldg - {building}"[:31]

            generate_site_report(
                bldg_df,
                sheet_name,
                workbook,
                date_cols,
                aggregate_floors=True
            )

    workbook.close()

def generate_site_report(df, sheet_name, workbook, date_cols, aggregate_floors=False):
    worksheet = workbook.add_worksheet(name=sheet_name[:31])

    worksheet.set_column('A:A', 20.5)
    worksheet.set_column(1, 35, 14.8)

    fmt = lambda **opts: workbook.add_format(opts)

    # ---------- Borders ----------
    # ---------- Borders ----------
    row_border = {
        'top': 1,
        'bottom': 1
    }

    day_sep_left = {
        'left': 2
    }

    BORDER_COLOR = '#808080'

    # ---------- Title / header formats ----------
    merge_format = fmt(
        align='center', valign='vcenter',
        fg_color='5C5B5A', font_color='white', font_size=14
    )

    label_format = fmt(
        align='center', valign='vcenter',
        fg_color='5C5B5A', font_color='white',
        font_size=12, text_wrap=1
    )

    header_format = fmt(
        align='center', valign='vcenter',
        fg_color='5C5B5A', font_color='white',
        font_size=10, bottom=2
    )

    # Header with vertical day separator (LEFT border)
    header_format_day_sep = fmt(
        align='center', valign='vcenter',
        fg_color='5C5B5A', font_color='white',
        font_size=10, bottom=2,
        left=2, border_color='#808080'
    )

    day_header_center = fmt(
        align='center', valign='vcenter',
        bold=1, font_size=10,
        bottom=2, left=2, border_color='#808080'
    )

    bottom_title = fmt(align='center', font_size=10, underline=1)
    bold_only = fmt(bold=1)

    # ---------- Row label formats ----------
    main_site_format = fmt(
        bold=1, bottom=1, font_size=10,
        bottom_color='#0000EE', align='right'
    )

    main_site_loc_format = fmt(
        bold=1, bottom=1, font_size=10,
        bottom_color='#0000EE', align='left'
    )

    sub_site_format = fmt(
        bottom=1, align='right',
        font_size=10, bottom_color='#800080'
    )

    sub_site_loc_format = fmt(
        bottom=1, align='left',
        font_size=10, bottom_color='#800080'
    )

    ssid_format = fmt(align='right', font_size=10, bg_color='#C0C0C0')
    ssid_name_format = fmt(align='left', bg_color='#C0C0C0', font_size=10)

    # ---------- Alternating day block formats ----------
    day1_sessions_fmt = fmt(
        bg_color='#F2F2F2', align='right',
        left=2,right=2, border_color='#808080', **row_border
    )
    day1_users_fmt = fmt(
        bg_color='#F2F2F2', align='right',
        **row_border
    )

    day2_sessions_fmt = fmt(
        bg_color='#FFFFFF', align='right',
        left=2,right=2, border_color='#808080', **row_border
    )
    day2_users_fmt = fmt(
        bg_color='#FFFFFF', align='right',
        **row_border
    )

    # ---------- Title ----------
    month = df['start_time'].dt.strftime('%B').mode()[0]
    worksheet.merge_range('A1:E1', f"WiFi Statistics Summary Report", merge_format)
    worksheet.merge_range('A2:A7', sheet_name, label_format)

    # ---------- Summary ----------
    worksheet.write('C4', 'Client User Summary', bold_only)
    worksheet.write('C5', 'Number of Sessions', bottom_title)
    worksheet.write('D5', 'Number of Users', bottom_title)
    worksheet.write('C6', len(df))
    worksheet.write('D6', df['client_mac'].nunique())

    # ---------- Static headers ----------
    worksheet.write('A8', 'Locations', header_format)
    worksheet.write('B8', 'SSID', header_format)
    worksheet.write('C8', 'Number of Sessions', header_format)
    worksheet.write('D8', 'Number of Users', header_format)
    worksheet.write('E8', '', header_format)

    # ---------- Day headers ----------
    for idx, day in enumerate(date_cols):
        base_col = 5 + (idx * 2)

        worksheet.merge_range(
            5, base_col, 5, base_col + 1,
            day.strftime('%d-%b'),
            day_header_center
        )

        worksheet.write(6, base_col,     'Sessions', header_format_day_sep)
        worksheet.write(6, base_col + 1, 'Users',    header_format)

        day_df = df[df['session_date'].dt.normalize() == day]

        sessions_fmt = day1_sessions_fmt if idx % 2 == 0 else day2_sessions_fmt
        users_fmt    = day1_users_fmt    if idx % 2 == 0 else day2_users_fmt

        worksheet.write(7, base_col,     len(day_df), sessions_fmt)
        worksheet.write(7, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

    # ---------- Time range ----------
    total_col = 5 + (len(date_cols) * 2)
    timeset = sorted(df['end_time'].tolist())

    worksheet.write(4, total_col + 2, 'Time Stamps from Client Summary', bold_only)
    worksheet.write(5, total_col + 2, 'Start time:')
    worksheet.write(5, total_col + 3, str(timeset[0]))
    worksheet.write(6, total_col + 2, 'End time:')
    worksheet.write(6, total_col + 3, str(timeset[-1]))

    # ---------- Data rows ----------
    cursor = 8

    for location in df['location'].unique():
        loc_df = df[df['location'] == location]
        cursor += 1

        worksheet.write(f'A{cursor}', f"    {location}", main_site_loc_format)
        worksheet.write(f'C{cursor}', len(loc_df), main_site_format)
        worksheet.write(f'D{cursor}', loc_df['client_mac'].nunique(), main_site_format)

        for i, day in enumerate(date_cols):
            base_col = 5 + (i * 2)
            day_df = loc_df[loc_df['session_date'].dt.normalize() == day]

            sessions_fmt = day1_sessions_fmt if i % 2 == 0 else day2_sessions_fmt
            users_fmt    = day1_users_fmt    if i % 2 == 0 else day2_users_fmt

            worksheet.write(cursor - 1, base_col,     len(day_df), sessions_fmt)
            worksheet.write(cursor - 1, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

        # SSID rows
        for ssid in loc_df['ssid'].unique():
            ssid_df = loc_df[loc_df['ssid'] == ssid]
            cursor += 1

            worksheet.write(f'B{cursor}', f"    {ssid}", ssid_name_format)
            worksheet.write(f'C{cursor}', len(ssid_df), ssid_format)
            worksheet.write(f'D{cursor}', ssid_df['client_mac'].nunique(), ssid_format)

            for i, day in enumerate(date_cols):
                base_col = 5 + (i * 2)
                day_df = ssid_df[ssid_df['session_date'].dt.normalize() == day]

                sessions_fmt = day1_sessions_fmt if i % 2 == 0 else day2_sessions_fmt
                users_fmt    = day1_users_fmt    if i % 2 == 0 else day2_users_fmt

                worksheet.write(cursor - 1, base_col,     len(day_df), sessions_fmt)
                worksheet.write(cursor - 1, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

        # Building / sublocation rows
        group_col = "building" if aggregate_floors else "sublocation"

        for name in loc_df[group_col].dropna().unique():
            sub_df = loc_df[loc_df[group_col] == name]
            cursor += 1

            worksheet.write(f'A{cursor}', f"        {name}", sub_site_loc_format)
            worksheet.write(f'C{cursor}', len(sub_df), sub_site_format)
            worksheet.write(f'D{cursor}', sub_df['client_mac'].nunique(), sub_site_format)

            for i, day in enumerate(date_cols):
                base_col = 5 + (i * 2)
                day_df = sub_df[sub_df['session_date'].dt.normalize() == day]

                sessions_fmt = day1_sessions_fmt if i % 2 == 0 else day2_sessions_fmt
                users_fmt    = day1_users_fmt    if i % 2 == 0 else day2_users_fmt

                worksheet.write(cursor - 1, base_col,     len(day_df), sessions_fmt)
                worksheet.write(cursor - 1, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

    for idx in range(len(date_cols)):
        base_col = 5 + (idx * 2)
        worksheet.set_column(base_col, base_col, 14.8)
        worksheet.set_column(base_col + 1, base_col + 1, 14.8)

    # # SSID pie chart
    # ssid_counts = df['ssid'].value_counts()
    # chart_row = cursor + 5
    # worksheet.merge_range(f'A{chart_row}:E{chart_row}', 'Unique Clients by SSID', merge_format)
    #
    # for i, (ssid, count) in enumerate(ssid_counts.items()):
    #     worksheet.write(chart_row + 1 + i, total_col + 7, ssid)
    #     worksheet.write(chart_row + 1 + i, total_col + 8, count)
    #
    # chart = workbook.add_chart({'type': 'pie'})
    # chart.add_series({
    #     'categories': [sheet_name[:31], chart_row + 1, total_col + 7, chart_row + len(ssid_counts), total_col + 7],
    #     'values':     [sheet_name[:31], chart_row + 1, total_col + 8, chart_row + len(ssid_counts), total_col + 8],
    # })
    # chart.set_style(10)
    # chart.set_size({'width': 540, 'height': 432})
    # worksheet.insert_chart(f'A{chart_row + 1}', chart, {'x_offset': 25, 'y_offset': 15})
