import os
from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from datetime import datetime, timedelta
# ─── Imports & existing code above ───


app = Flask(__name__)
app.secret_key = 'replace-with-a-secure-random-key'

# Paths
DATA_DIR      = 'data'
RESOURCE_FILE = os.path.join(DATA_DIR, 'ResourceSheet.xlsx')
WORKITEM_FILE = os.path.join(DATA_DIR, 'WorkItem.xlsx')

# Default hours per day
DEF_WORKING = 6
DEF_YELLOW  = 3
DEF_RED     = 0

def ensure_files():
    """Create Excel files with required sheets if they don't exist."""
    os.makedirs(DATA_DIR, exist_ok=True)

    if not os.path.exists(RESOURCE_FILE):
        df_res = pd.DataFrame(columns=[
            'ResourceId', 'ResourceName', 'WorkingHrs'
        ])
        df_timeoff = pd.DataFrame(columns=['ResourceId', 'TimeOffDate', 'WorkingHrs'])
        df_hol = pd.DataFrame(columns=['HolidayDate'])
        with pd.ExcelWriter(RESOURCE_FILE, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name='Resource', index=False)
            df_timeoff.to_excel(writer, sheet_name='TimeOff', index=False)
            df_hol.to_excel(writer, sheet_name='Holiday', index=False)

    if not os.path.exists(WORKITEM_FILE):
        df_wi = pd.DataFrame(columns=[
            'WorkId', 'ProjectName', 'Estimate',
            'ProjStart', 'ProjEnd',
            'AssignedResource', 'AssignDatetime', 'Status',
            'RemainingHours', 'Notes'
        ])
        df_wi.to_excel(WORKITEM_FILE, index=False)

def load_resources():
    """Read the Resource sheet with ResourceId as string.

    Some existing spreadsheets may lack the ``WorkingHrs`` column. Ensure it is
    present and fill missing values with ``DEF_WORKING`` so callers can rely on
    it.
    """
    df = pd.read_excel(
        RESOURCE_FILE,
        sheet_name='Resource',
        dtype={'ResourceId': str, 'ResourceName': str},
    )
    if 'WorkingHrs' not in df.columns:
        df['WorkingHrs'] = DEF_WORKING
    else:
        df['WorkingHrs'] = df['WorkingHrs'].fillna(DEF_WORKING).astype(int)
    return df
# … your ensure_files(), load_resources(), load_holidays(), save_resources() etc.

def load_timeoff():
    """Return TimeOff DataFrame with normalized types.

    ``WorkingHrs`` represents the number of working hours on the time‑off day
    (defaulting to 0 meaning a full day off).
    """
    df = pd.read_excel(
        RESOURCE_FILE,
        sheet_name='TimeOff',
        dtype={'ResourceId': str, 'TimeOffDate': object, 'WorkingHrs': float}
    )
    # Keep only known columns to avoid duplicate names when merging later on.
    df = df.loc[:, [c for c in ['ResourceId', 'TimeOffDate', 'WorkingHrs'] if c in df.columns]]

    if 'TimeOffDate' in df.columns:
        df['TimeOffDate'] = (
            pd.to_datetime(df['TimeOffDate'], errors='coerce')
              .dt.date
        )
        df = df.dropna(subset=['TimeOffDate'])

    if 'WorkingHrs' not in df.columns:
        df['WorkingHrs'] = 0.0
    else:
        df['WorkingHrs'] = pd.to_numeric(df['WorkingHrs'], errors='coerce').fillna(0)

    return df

@app.route('/timeoff')
def view_timeoff():
    df_timeoff = load_timeoff()
    df_res = load_resources()

    today = datetime.today().date()
    default_start = today.isoformat()
    default_end = today.replace(month=12, day=31).isoformat()

    resource_filter = request.args.get('resource', '').strip()
    start = request.args.get('start', default_start).strip() or default_start
    end = request.args.get('end', default_end).strip() or default_end

    if resource_filter:
        df_timeoff = df_timeoff[df_timeoff['ResourceId'] == resource_filter]

    # Merge to pull in ResourceName from the Resource sheet
    df = df_timeoff.merge(
        df_res[['ResourceId', 'ResourceName']],
        on='ResourceId',
        how='left'
    )
    df = df[['ResourceId', 'ResourceName', 'WorkingHrs', 'TimeOffDate']]
    df = df.sort_values(['ResourceId', 'TimeOffDate'])

    available = total = timeoff_hours = None
    if resource_filter and start and end:
        try:
            start_dt = datetime.fromisoformat(start)
            end_dt = datetime.fromisoformat(end)
            available, total, timeoff_hours = compute_work_hours(
                resource_filter, start_dt, end_dt, df_res
            )
        except ValueError:
            flash('Invalid date range.', 'error')

    return render_template(
        'timeoff.html',
        timeoff=df.to_dict('records'),
        resources=df_res.to_dict('records'),
        selected_resource=resource_filter,
        start=start,
        end=end,
        available_hours=available,
        total_hours=total,
        timeoff_hours=timeoff_hours,
    )

@app.route('/add_timeoff', methods=['GET','POST'])
def add_timeoff():
    df_res = load_resources()
    df_timeoff = load_timeoff()
    df_hol = pd.DataFrame({'HolidayDate': load_holidays()})

    if request.method == 'POST':
        rid  = request.form['resource_id']
        date = request.form['timeoff_date']
        work_val = request.form.get('working_hrs', '').strip()

        try:
            timeoff_date = datetime.fromisoformat(date).date()
        except Exception:
            flash("Invalid date format.", 'error')
            return redirect(url_for('add_timeoff'))

        try:
            work_hrs = float(work_val) if work_val else 0
        except ValueError:
            flash("Invalid working hours.", 'error')
            return redirect(url_for('add_timeoff'))

        # append new TimeOff row
        new = {
            'ResourceId': rid,
            'TimeOffDate': timeoff_date,
            'WorkingHrs': work_hrs
        }
        df_timeoff = pd.concat([df_timeoff, pd.DataFrame([new])], ignore_index=True)

        # save back
        save_resources(df_res, df_timeoff, df_hol)
        flash(f"TimeOff added for {rid} on {timeoff_date}", 'success')
        return redirect(url_for('view_timeoff'))

    # GET: show form
    resources = df_res.to_dict('records')
    return render_template('add_timeoff.html', resources=resources)

def load_holidays():
    """Read the Holiday sheet and return a list of dates."""
    df = pd.read_excel(
        RESOURCE_FILE,
        sheet_name='Holiday',
        dtype={'HolidayDate': object}
    )
    if 'HolidayDate' in df:
        return (
            pd.to_datetime(df['HolidayDate'], errors='coerce')
              .dt.date
              .dropna()
              .tolist()
        )
    return []

def load_workitems():
    """Read the WorkItem file with proper datetime parsing."""
    df = pd.read_excel(
        WORKITEM_FILE,
        parse_dates=['ProjStart', 'ProjEnd', 'AssignDatetime'],
        dtype={'AssignedResource': str}
    )
    if 'RemainingHours' not in df.columns:
        df['RemainingHours'] = df['Estimate']
    if 'Notes' not in df.columns:
        df['Notes'] = ''
    return df

def save_resources(df_res, df_timeoff, df_hol):
    """Overwrite the ResourceSheet.xlsx with updated sheets."""
    with pd.ExcelWriter(RESOURCE_FILE, engine='openpyxl', mode='w') as writer:
        df_res.to_excel(writer, sheet_name='Resource', index=False)
        df_timeoff.to_excel(writer, sheet_name='TimeOff', index=False)
        df_hol.to_excel(writer, sheet_name='Holiday',  index=False)

def save_workitems(df_wi):
    """Overwrite the WorkItem.xlsx."""
    df_wi.to_excel(WORKITEM_FILE, index=False)

def compute_work_hours(resource_id, start_dt, end_dt, df_res):
    """
    Calculate working hours for a resource between two datetimes.

    Returns a tuple ``(available, total, timeoff)`` where:

    * ``total``   – working hours for business days in the range, excluding
      holidays.
    * ``timeoff`` – working hours removed due to time‑off entries that fall on
      those business days.
    * ``available`` – the remaining hours after deducting ``timeoff`` from
      ``total``.
    """

    timeoff_df = load_timeoff()
    hols = set(load_holidays())
    row = df_res.set_index('ResourceId').loc[resource_id]
    workinghrs = int(row['WorkingHrs'])

    # Map of date -> minimum working hours remaining on that day
    timeoff_map = (
        timeoff_df[timeoff_df['ResourceId'] == resource_id]
        .groupby('TimeOffDate')['WorkingHrs']
        .min()
        .to_dict()
    )

    total = 0
    timeoff_hours = 0
    curr = start_dt.date()
    end = end_dt.date()
    while curr <= end:
        # business day check
        if curr.weekday() < 5 and curr not in hols:
            total += workinghrs
            if curr in timeoff_map:
                w = max(0, min(workinghrs, timeoff_map[curr]))
                timeoff_hours += workinghrs - w
        curr += timedelta(days=1)

    available = total - timeoff_hours
    return available, total, timeoff_hours


def available_hours(resource_id, start_dt, end_dt, df_res):
    """Backwards compatible helper returning only available hours."""
    available, _total, _timeoff = compute_work_hours(
        resource_id, start_dt, end_dt, df_res
    )
    return available

def assess_status(avail, required):
    """
    Determine status:
    - Available: avail >= required
    - At Risk:   avail >= 70% of required
    - Insufficient otherwise
    """
    if avail >= required:
        return 'OnTrack'
    if avail >= required * 0.7:
        return 'AtRisk'
    return 'OffTrack'

# Initialize files on startup
ensure_files()

@app.route('/')
def index():
    """Default home redirects to the work item list."""
    return redirect(url_for('view_workitems'))

@app.route('/resources')
def view_resources():
    df = load_resources()
    return render_template('resources.html', resources=df.to_dict('records'))

@app.route('/add_resource', methods=['GET','POST'])
def add_resource():
    if request.method == 'POST':
        df_res = load_resources()
        df_timeoff = load_timeoff()
        df_hol = pd.DataFrame({'HolidayDate': load_holidays()})

        new = {
            'ResourceId':   request.form['res_id'].strip(),
            'ResourceName': request.form['res_name'].strip(),
            'WorkingHrs':    int(request.form.get('workinghrs') or DEF_WORKING)
        }
        df_res = pd.concat([df_res, pd.DataFrame([new])], ignore_index=True)
        save_resources(df_res, df_timeoff, df_hol)
        flash(f"Resource {new['ResourceId']} added.", 'success')
        return redirect(url_for('view_resources'))

    return render_template('add_resource.html')

@app.route('/holidays')
def view_holidays():
    hols = load_holidays()
    return render_template('holidays.html', holidays=hols)


@app.route('/add_holiday', methods=['GET', 'POST'])
def add_holiday():
    df_res = load_resources()
    df_timeoff = load_timeoff()
    df_hol = pd.DataFrame({'HolidayDate': load_holidays()})

    if request.method == 'POST':
        date = request.form['holiday_date']
        try:
            hol_date = datetime.fromisoformat(date).date()
        except Exception:
            flash('Invalid date format.', 'error')
            return redirect(url_for('add_holiday'))

        if hol_date in df_hol['HolidayDate'].tolist():
            flash('Holiday already exists.', 'warning')
        else:
            df_hol = pd.concat([
                df_hol,
                pd.DataFrame([{'HolidayDate': hol_date}])
            ], ignore_index=True)
            save_resources(df_res, df_timeoff, df_hol)
            flash(f'Holiday {hol_date} added.', 'success')
        return redirect(url_for('view_holidays'))

    return render_template('add_holiday.html')


@app.route('/project', methods=['GET','POST'])
def project_form():
    if request.method == 'POST':
        name     = request.form.get('name','').strip()
        try:
            estimate = float(request.form['estimate'])
        except ValueError:
            flash("Estimate must be a number.", 'error')
            return redirect(url_for('project_form'))

        ps = request.form['proj_start']
        pe = request.form['proj_end']
        try:
            ps_dt = datetime.fromisoformat(ps)
            pe_dt = datetime.fromisoformat(pe)
        except Exception:
            flash("Invalid date/time format.", 'error')
            return redirect(url_for('project_form'))

        df_res = load_resources()
        if df_res.empty:
            flash("Add at least one resource first.", 'warning')
            return redirect(url_for('add_resource'))

        statuses = []
        for _, r in df_res.iterrows():
            rid   = r['ResourceId']
            avail = available_hours(rid, ps_dt, pe_dt, df_res)
            stat  = assess_status(avail, estimate)
            statuses.append({
                'ResourceId':     rid,
                'ResourceName':   r['ResourceName'],
                'AvailableHours': avail,
                'Status':         stat
            })

        return render_template(
            'select_resource.html',
            project_name=name,
            estimate=estimate,
            proj_start=ps,
            proj_end=pe,
            statuses=statuses
        )

    today = datetime.today().date().isoformat()
    return render_template('project_form.html', today=today)

@app.route('/assign', methods=['POST'])
def assign():
    df_wi = load_workitems()
    wi_id = len(df_wi) + 1

    name       = request.form.get('project_name') or f'WorkItem#{wi_id}'
    estimate   = float(request.form['estimate'])
    proj_start = datetime.fromisoformat(request.form['proj_start'])
    proj_end   = datetime.fromisoformat(request.form['proj_end'])
    rid        = request.form['resource_id']
    status     = request.form['status']
    assign_dt  = datetime.now()

    new = {
        'WorkId':           wi_id,
        'ProjectName':      name,
        'Estimate':         estimate,
        'ProjStart':        proj_start,
        'ProjEnd':          proj_end,
        'AssignedResource': rid,
        'AssignDatetime':   assign_dt,
        'Status':           status,
        'RemainingHours':   estimate
    }

    df_wi = pd.concat([df_wi, pd.DataFrame([new])], ignore_index=True)
    save_workitems(df_wi)
    flash(f"Assigned {rid} to {name}.", 'success')
    return redirect(url_for('view_workitems'))


@app.route('/edit_workitem/<int:work_id>', methods=['GET', 'POST'])
def edit_workitem(work_id):
    df_wi = load_workitems()
    if work_id not in df_wi['WorkId'].values:
        flash('WorkItem not found.', 'error')
        return redirect(url_for('view_workitems'))

    idx = df_wi.index[df_wi['WorkId'] == work_id][0]
    row = df_wi.loc[idx]

    if request.method == 'POST':
        try:
            remaining = float(request.form['remaining_hours'])
        except ValueError:
            flash('Remaining hours must be a number.', 'error')
            return redirect(url_for('edit_workitem', work_id=work_id))

        df_wi.at[idx, 'RemainingHours'] = remaining
        df_wi.at[idx, 'Notes'] = request.form.get('notes', row.get('Notes', ''))

        if df_wi.at[idx, 'Status'] not in ('Paused', 'Completed'):
            df_res = load_resources()
            avail = available_hours(
                row['AssignedResource'], datetime.now(), row['ProjEnd'], df_res
            )
            df_wi.at[idx, 'Status'] = assess_status(avail, remaining)

        save_workitems(df_wi)
        flash('WorkItem updated.', 'success')
        return redirect(url_for('view_workitems'))

    return render_template('edit_workitem.html', item=row)

@app.route('/update_notes/<int:work_id>', methods=['POST'])
def update_notes(work_id):
    df_wi = load_workitems()
    if work_id not in df_wi['WorkId'].values:
        flash('WorkItem not found.', 'error')
        return redirect(url_for('view_workitems'))
    
    idx = df_wi.index[df_wi['WorkId'] == work_id][0]

    new_note = request.form['new_note']
    timestamp = datetime.now().strftime('%Y-%m-%d')
    existing_notes = df_wi.at[idx, 'Notes']
    new_entry = f"{timestamp}: {new_note}"
    df_wi.at[idx, 'Notes'] = f"{existing_notes}\n{new_entry}" if pd.notna(existing_notes) and existing_notes else new_entry
    
    save_workitems(df_wi)
    flash('Note added.', 'success')
    return redirect(url_for('view_workitems'))


@app.route('/pause_workitem/<int:work_id>')
def pause_workitem(work_id):
    df_wi = load_workitems()
    if work_id in df_wi['WorkId'].values:
        idx = df_wi.index[df_wi['WorkId'] == work_id][0]
        df_wi.at[idx, 'Status'] = 'Paused'
        save_workitems(df_wi)
        flash('WorkItem paused.', 'success')
    else:
        flash('WorkItem not found.', 'error')
    return redirect(url_for('view_workitems'))


@app.route('/complete_workitem/<int:work_id>')
def complete_workitem(work_id):
    df_wi = load_workitems()
    if work_id in df_wi['WorkId'].values:
        idx = df_wi.index[df_wi['WorkId'] == work_id][0]
        df_wi.at[idx, 'Status'] = 'Completed'
        df_wi.at[idx, 'RemainingHours'] = 0
        save_workitems(df_wi)
        flash('WorkItem completed.', 'success')
    else:
        flash('WorkItem not found.', 'error')
    return redirect(url_for('view_workitems'))


@app.route('/delete_workitem/<int:work_id>')
def delete_workitem(work_id):
    df_wi = load_workitems()
    if work_id in df_wi['WorkId'].values:
        df_wi = df_wi[df_wi['WorkId'] != work_id]
        save_workitems(df_wi)
        flash('WorkItem deleted.', 'success')
    else:
        flash('WorkItem not found.', 'error')
    return redirect(url_for('view_workitems'))

def _render_workitems(status_filter=None):
    """Helper to render work items with optional filtering and sorting."""
    df = load_workitems()
    df_res = load_resources()[['ResourceId', 'ResourceName']]

    # merge so we can display ResourceName
    df = df.merge(
        df_res,
        left_on='AssignedResource',
        right_on='ResourceId',
        how='left'
    )

    search = request.args.get('search', '').strip()
    if search:
        df = df[df['ProjectName'].str.contains(search, case=False, na=False)]

    resource_filter = request.args.get('resource', '').strip()
    if resource_filter:
        df = df[df['AssignedResource'] == resource_filter]

    if status_filter:
        df = df[df['Status'] == status_filter]

    sort_key = request.args.get('sort', '').strip()
    if sort_key == 'start':
        df = df.sort_values('ProjStart')
    elif sort_key == 'end':
        df = df.sort_values('ProjEnd')
    elif sort_key == 'resource':
        df = df.sort_values('ResourceName')

    # compute status based on remaining hours and availability
    df_res_all = load_resources()
    def _calc_status(row):
        if row['Status'] in ('Paused', 'Completed'):
            return row['Status']
        avail = available_hours(row['AssignedResource'], datetime.now(), row['ProjEnd'], df_res_all)
        return assess_status(avail, row['RemainingHours'])

    if not df.empty:
        df['Status'] = df.apply(_calc_status, axis=1)

    # format dates for display
    if not df.empty:
        df['ProjStart']      = df['ProjStart'].dt.strftime('%Y-%m-%d')
        df['ProjEnd']        = df['ProjEnd'].dt.strftime('%Y-%m-%d')
        df['AssignDatetime'] = df['AssignDatetime'].dt.strftime('%Y-%m-%d %H:%M')

    resources = df_res.to_dict('records')
    return render_template(
        'workitems.html',
        workitems=df.to_dict('records'),
        resources=resources,
        search=search,
        selected_resource=resource_filter,
        sort=sort_key,
        status=status_filter
    )


@app.route('/workitems')
def view_workitems():
    """Show all work items."""
    return _render_workitems()


@app.route('/completed_workitems')
def view_completed_workitems():
    """Show completed work items."""
    return _render_workitems(status_filter='Completed')


@app.route('/paused_workitems')
def view_paused_workitems():
    """Show paused work items."""
    return _render_workitems(status_filter='Paused')

if __name__ == '__main__':
    app.run(debug=True)
