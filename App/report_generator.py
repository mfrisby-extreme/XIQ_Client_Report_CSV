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

def generate_excel_report(data, selected_sites, output_path, date_from=None, date_to=None, aggregate_floors=False, tab_per_building=False):

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

    workbook = xlsxwriter.Workbook(output_path)

    # Aggregate report (multiple sites)
    if len(selected_sites) > 1:
        agg_df = df[df['location'].isin(selected_sites)]
        # FIX #3: compute date_cols scoped to this data slice
        agg_date_cols = sorted(agg_df['session_date'].dt.normalize().unique())
        generate_site_report(
            agg_df,
            "Report",
            workbook,
            agg_date_cols,
            aggregate_floors=aggregate_floors,
            is_aggregate=True  # FIX #4: flag so roamer logic can use cross-site scope
        )

    # Per-site reports
    for site in selected_sites:
        site_df = df[df['location'] == site]
        if not site_df.empty:
            # FIX #3: date_cols scoped to this site only
            site_date_cols = sorted(site_df['session_date'].dt.normalize().unique())
            generate_site_report(
                site_df,
                site,
                workbook,
                site_date_cols,
                aggregate_floors=aggregate_floors,
                is_aggregate=False
            )

    # Per-building reports (ONLY when aggregating floors)
    if aggregate_floors and tab_per_building:
        building_df = df[df['location'].isin(selected_sites)]

        for building in sorted(building_df['building'].dropna().unique()):
            bldg_df = building_df[building_df['building'] == building]

            if bldg_df.empty:
                continue

            sheet_name = f"Bldg - {building}"[:31]
            bldg_date_cols = sorted(bldg_df['session_date'].dt.normalize().unique())

            generate_site_report(
                bldg_df,
                sheet_name,
                workbook,
                bldg_date_cols,
                aggregate_floors=True,
                is_aggregate=False
            )

    workbook.close()

def generate_site_report(df, sheet_name, workbook, date_cols, aggregate_floors=False, is_aggregate=False):
    worksheet = workbook.add_worksheet(name=sheet_name[:31])

    worksheet.set_column('A:A', 30)
    worksheet.set_column(1, 35, 14.8)
    worksheet.set_column('C:C', 30)  # Number of Sessions
    worksheet.set_column('D:D', 20)  # Number of Unique Users
    worksheet.set_column('E:E', 20)  # Number of Users

    fmt = lambda **opts: workbook.add_format(opts)

    row_border = {
        'top': 1,
        'bottom': 1
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
        left=2, right=2, border_color='#808080', **row_border
    )
    day1_users_fmt = fmt(
        bg_color='#F2F2F2', align='right',
        **row_border
    )

    day2_sessions_fmt = fmt(
        bg_color='#FFFFFF', align='right',
        left=2, right=2, border_color='#808080', **row_border
    )
    day2_users_fmt = fmt(
        bg_color='#FFFFFF', align='right',
        **row_border
    )

    # ---------- Title ----------
    month = df['start_time'].dt.strftime('%B').mode()[0]
    worksheet.merge_range('A1:E1', "WiFi Statistics Summary Report", merge_format)
    # Label column spans rows 2-9 to accommodate the expanded summary block
    worksheet.merge_range('A2:A9', sheet_name, label_format)

    # ---------- Summary ----------
    worksheet.write('C4', 'Client User Summary', bold_only)
    worksheet.write('C5', 'Number of Sessions', bottom_title)
    worksheet.write('D5', 'Number of Unique Users', bottom_title)
    worksheet.write_comment('D5',
                            'This is the total number of unique mac addresses over the time period. '
                            'A single user/device is counted only once regardless of how many '
                            'times/days they visited.',
                            {'x_scale': 2, 'y_scale': 1.5}
                            )
    worksheet.write('E5', 'Number of Users', bottom_title)
    worksheet.write_comment('E5',
                            'This is the sum of each day\'s unique user count with no cross-day '
                            'deduplication. A user who visits on 3 different days is counted 3 times here, ',
                            {'x_scale': 2.5, 'y_scale': 2}
                            )

    worksheet.write('C6', len(df))
    worksheet.write('D6', df['client_mac'].nunique())
    # Sum of per-day unique user counts — no cross-day deduplication.
    # This matches what the daily columns add up to, explaining the gap vs column D.
    daily_user_sum = int(
        df.groupby(df['session_date'].dt.normalize())['client_mac'].nunique().sum()
    )
    worksheet.write('E6', daily_user_sum)

    group_col = "building" if aggregate_floors else "sublocation"

    # FIX #4: On the aggregate tab, count roamers across site+building combinations
    # to catch users who roamed across different sites, not just within one site.
    # On per-site tabs, count roamers across sub-locations/buildings within the site.
    if is_aggregate:
        df['_roamer_key'] = df['location'] + ' | ' + df[group_col].astype(str)
        visits_per_client = df.groupby("client_mac")['_roamer_key'].nunique()
        df.drop(columns=['_roamer_key'], inplace=True)
        roamer_label      = "Users Visiting Multiple Sites/Buildings"
        group_label       = "sites/buildings"
    else:
        visits_per_client = df.groupby("client_mac")[group_col].nunique()
        roamer_label      = "Users Visiting Multiple Buildings" if aggregate_floors else "Users Visiting Multiple Sublocations"
        group_label       = "buildings" if aggregate_floors else "sublocations"

    multi_loc_clients = int((visits_per_client > 1).sum())
    # Extra appearances = sum of (groups_visited - 1) per roamer.
    # This is the number that reconciles the totals:
    #   sum(per-group unique users) = total unique users + extra appearances
    extra_appearances = int((visits_per_client[visits_per_client > 1] - 1).sum())

    note_fmt = fmt(italic=1, font_size=9, font_color='#444444')

    worksheet.write("C7",  roamer_label, bold_only)
    worksheet.write_comment('C7',
                            'The number of unique devices (MAC addresses) that were seen at more than one '
                            'building during this reporting period. For example, if 310 is shown here, '
                            'it means 310 individual devices connected at two or more different buildings. '
                            'These devices are counted once in the global "Number of Unique Users" total, '
                            'but appear in each building\'s individual user count.',
                            {'x_scale': 2.5, 'y_scale': 2}
                            )
    worksheet.write("D7",  multi_loc_clients)
    worksheet.write("C8",  f"Extra Appearances", bold_only)
    worksheet.write_comment('C8',
                            'The total number of additional building appearances generated by roaming devices. '
                            'A device visiting 2 buildings adds 1 extra appearance; visiting 3 buildings adds 2, and so on. '
                            'This is the reconciliation number: '
                            'Sum of per-building user counts = Total Unique Users + Extra Appearances. '
                            'For example, if Total Unique Users is 1000 and Extra Appearances is 200, '
                            'the per-building counts will sum to exactly 1200.',
                            {'x_scale': 3, 'y_scale': 2.5}
                            )
    worksheet.write("D8",  extra_appearances)
    worksheet.write("C9",  "Sum of per-group users = Total Users + Extra Appearances", note_fmt)
    worksheet.write("D9",  f"{df['client_mac'].nunique()} + {extra_appearances} = {df['client_mac'].nunique() + extra_appearances}", note_fmt)

    # ---------- Static headers ----------
    # Summary now ends at row 9; column headers on row 11, day-totals on row 12.
    HEADER_ROW   = 10   # 0-indexed (Excel row 11)
    TOTAL_ROW    = 11   # 0-indexed (Excel row 12)
    worksheet.write(HEADER_ROW, 0, 'Locations',                        header_format)
    worksheet.write(HEADER_ROW, 1, 'SSID',                             header_format)
    worksheet.write(HEADER_ROW, 2, 'Number of Sessions',               header_format)
    worksheet.write(HEADER_ROW, 3, 'Number of Unique Users',           header_format)
    worksheet.write(HEADER_ROW, 4, 'Number of Users',  header_format)

    # ---------- Day headers ----------
    for idx, day in enumerate(date_cols):
        base_col = 5 + (idx * 2)

        # Day-name label merges above the Sessions/Users row
        worksheet.merge_range(
            HEADER_ROW - 1, base_col, HEADER_ROW - 1, base_col + 1,
            day.strftime('%d-%b'),
            day_header_center
        )
        # Sessions/Users sub-labels on the header row
        worksheet.write(HEADER_ROW, base_col,     'Sessions', header_format_day_sep)
        worksheet.write(HEADER_ROW, base_col + 1, 'Users',    header_format)

        day_df = df[df['session_date'].dt.normalize() == day]
        sessions_fmt = day1_sessions_fmt if idx % 2 == 0 else day2_sessions_fmt
        users_fmt    = day1_users_fmt    if idx % 2 == 0 else day2_users_fmt

        worksheet.write(TOTAL_ROW, base_col,     len(day_df),                    sessions_fmt)
        worksheet.write(TOTAL_ROW, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

    # Column E total row: sum of daily unique users across all days
    total_daily_sum = int(
        df.groupby(df['session_date'].dt.normalize())['client_mac'].nunique().sum()
    )
    worksheet.write(TOTAL_ROW, 4, total_daily_sum, day1_users_fmt)

    # ---------- Time range ----------
    total_col = 5 + (len(date_cols) * 2)
    timeset = sorted(df['end_time'].tolist())

    worksheet.write(4, total_col + 2, 'Time Stamps from Client Summary', bold_only)
    worksheet.write(5, total_col + 2, 'Start time:')
    worksheet.write(5, total_col + 3, str(timeset[0]))
    worksheet.write(6, total_col + 2, 'End time:')
    worksheet.write(6, total_col + 3, str(timeset[-1]))

    # ---------- Data rows ----------
    # First data row is Excel row 13 (1-indexed 13 = cursor starts at 12,
    # then cursor += 1 before each write brings it to 13).
    cursor = 12
    group_rows = []  # will hold the 1-indexed Excel rows for each building/sublocation

    for location in df['location'].unique():
        loc_df = df[df['location'] == location]
        cursor += 1

        worksheet.write(f'A{cursor}', f"    {location}", main_site_loc_format)
        worksheet.write(f'C{cursor}', len(loc_df), main_site_format)
        worksheet.write(f'D{cursor}', loc_df['client_mac'].nunique(), main_site_format)
        worksheet.write(f'E{cursor}', int(loc_df.groupby(loc_df['session_date'].dt.normalize())['client_mac'].nunique().sum()), main_site_format)

        for i, day in enumerate(date_cols):
            base_col = 5 + (i * 2)
            day_df = loc_df[loc_df['session_date'].dt.normalize() == day]

            sessions_fmt = day1_sessions_fmt if i % 2 == 0 else day2_sessions_fmt
            users_fmt    = day1_users_fmt    if i % 2 == 0 else day2_users_fmt

            # FIX #2: use cursor - 1 consistently (0-indexed = 1-indexed cursor minus 1)
            worksheet.write(cursor - 1, base_col,     len(day_df), sessions_fmt)
            worksheet.write(cursor - 1, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

        # SSID rows
        for ssid in loc_df['ssid'].unique():
            ssid_df = loc_df[loc_df['ssid'] == ssid]
            cursor += 1

            worksheet.write(f'B{cursor}', f"    {ssid}", ssid_name_format)
            worksheet.write(f'C{cursor}', len(ssid_df), ssid_format)
            worksheet.write(f'D{cursor}', ssid_df['client_mac'].nunique(), ssid_format)
            worksheet.write(f'E{cursor}', int(ssid_df.groupby(ssid_df['session_date'].dt.normalize())['client_mac'].nunique().sum()), ssid_format)

            for i, day in enumerate(date_cols):
                base_col = 5 + (i * 2)
                day_df = ssid_df[ssid_df['session_date'].dt.normalize() == day]

                sessions_fmt = day1_sessions_fmt if i % 2 == 0 else day2_sessions_fmt
                users_fmt    = day1_users_fmt    if i % 2 == 0 else day2_users_fmt

                worksheet.write(cursor - 1, base_col,     len(day_df), sessions_fmt)
                worksheet.write(cursor - 1, base_col + 1, day_df['client_mac'].nunique(), users_fmt)

        # Building / sublocation rows
        for name in loc_df[group_col].dropna().unique():
            sub_df = loc_df[loc_df[group_col] == name]
            cursor += 1
            group_rows.append(cursor)  # track this row for the chart

            worksheet.write(f'A{cursor}', f"        {name}", sub_site_loc_format)
            worksheet.write(f'C{cursor}', len(sub_df), sub_site_format)
            worksheet.write(f'D{cursor}', sub_df['client_mac'].nunique(), sub_site_format)
            worksheet.write(f'E{cursor}', int(sub_df.groupby(sub_df['session_date'].dt.normalize())['client_mac'].nunique().sum()), sub_site_format)

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

        # ---------- Bar chart (summary/aggregate tabs only) ----------
        # Skip chart on per-building sub-tabs (sheet names starting with "Bldg - ")
    if sheet_name.startswith("Bldg - ") or not group_rows:
        return

    sname = sheet_name[:31]
    n_groups = len(group_rows)
    first_row = group_rows[0] - 1  # 0-indexed
    last_row = group_rows[-1] - 1  # 0-indexed

    chart = workbook.add_chart({'type': 'bar'})  # 'bar' = horizontal in xlsxwriter

    chart.add_series({
        'name': 'Number of Sessions',
        'categories': [sname, first_row, 0, last_row, 0],  # col A = location names
        'values': [sname, first_row, 2, last_row, 2],  # col C = sessions
        'fill': {'color': '#4FC3C8'},
        'data_labels': {'value': True, 'font': {'size': 8}},
    })
    chart.add_series({
        'name': 'Number of Users',
        'categories': [sname, first_row, 0, last_row, 0],  # col A = location names
        'values': [sname, first_row, 4, last_row, 4],  # col E = daily sum
        'fill': {'color': '#A8D5B5'},
        'data_labels': {'value': True, 'font': {'size': 8}},
    })
    chart.add_series({
        'name': 'Number of Unique Users',
        'categories': [sname, first_row, 0, last_row, 0],  # col A = location names
        'values': [sname, first_row, 3, last_row, 3],  # col D = unique users
        'fill': {'color': '#1A6B7C'},
        'data_labels': {'value': True, 'font': {'size': 8}},
    })


    group_label_title = "Building" if aggregate_floors else "Sublocation"
    chart.set_title({'name': f'{sname} — Session & User Count Chart'})
    chart.set_x_axis({'name': 'Count', 'major_gridlines': {'visible': True}})
    chart.set_y_axis({'name': group_label_title, 'reverse': True})
    chart.set_legend({'position': 'top'})
    chart.set_style(2)

    chart_height = min(max(300, n_groups * 40 + 100), 900)
    chart.set_size({'width': 900, 'height': chart_height})

    chart_anchor_row = cursor + 2
    worksheet.insert_chart(f'A{chart_anchor_row}', chart, {'x_offset': 5, 'y_offset': 5})