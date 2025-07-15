import os
from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = 'replace-with-a-secure-random-key'

# Paths
DATA_DIR      = 'data'
RESOURCE_FILE = os.path.join(DATA_DIR, 'ResourceSheet.xlsx')
WORKITEM_FILE = os.path.join(DATA_DIR, 'WorkItem.xlsx')

# Default hours per day
DEF_GREEN  = 6
DEF_YELLOW = 3
DEF_RED    = 0

def ensure_files():
    """Create Excel files with required sheets if they don't exist."""
    os.makedirs(DATA_DIR, exist_ok=True)

    if not os.path.exists(RESOURCE_FILE):
        df_res = pd.DataFrame(columns=[
            'ResourceId', 'ResourceName', 'GreenTime', 'YellowTime', 'RedTime'
        ])
        df_pto = pd.DataFrame(columns=['ResourceId', 'PTODate'])
        df_hol = pd.DataFrame(columns=['HolidayDate'])
        with pd.ExcelWriter(RESOURCE_FILE, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name='Resource', index=False)
            df_pto.to_excel(writer, sheet_name='PTO', index=False)
            df_hol.to_excel(writer, sheet_name='Holiday', index=False)

    if not os.path.exists(WORKITEM_FILE):
        df_wi = pd.DataFrame(columns=[
            'WorkId', 'ProjectName', 'Estimate',
            'ProjStart', 'ProjEnd',
            'AssignedResource', 'AssignDatetime', 'Status'
        ])
        df_wi.to_excel(WORKITEM_FILE, index=False)

def load_resources():
    """Read the Resource sheet with ResourceId as string."""
    return pd.read_excel(
        RESOURCE_FILE,
        sheet_name='Resource',
        dtype={'ResourceId': str, 'ResourceName': str,
               'GreenTime': int, 'YellowTime': int, 'RedTime': int}
    )

def load_pto():
    """Read the PTO sheet and coerce PTODate to date."""
    df = pd.read_excel(
        RESOURCE_FILE,
        sheet_name='PTO',
        dtype={'ResourceId': str, 'PTODate': object}
    )
    if 'PTODate' in df:
        df['PTODate'] = pd.to_datetime(
            df['PTODate'], errors='coerce'
        ).dt.date
    return df.dropna(subset=['PTODate'])

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
    return pd.read_excel(
        WORKITEM_FILE,
        parse_dates=['ProjStart', 'ProjEnd', 'AssignDatetime'],
        dtype={'AssignedResource': str}
    )

def save_resources(df_res, df_pto, df_hol):
    """Overwrite the ResourceSheet.xlsx with updated sheets."""
    with pd.ExcelWriter(RESOURCE_FILE, engine='openpyxl', mode='w') as writer:
        df_res.to_excel(writer, sheet_name='Resource', index=False)
        df_pto.to_excel(writer, sheet_name='PTO',      index=False)
        df_hol.to_excel(writer, sheet_name='Holiday',  index=False)

def save_workitems(df_wi):
    """Overwrite the WorkItem.xlsx."""
    df_wi.to_excel(WORKITEM_FILE, index=False)

def available_hours(resource_id, start_dt, end_dt, df_res):
    """
    Sum available hours for a resource between two dates,
    skipping PTO and holidays.
    """
    pto_df = load_pto()
    hols   = load_holidays()
    row    = df_res.set_index('ResourceId').loc[resource_id]
    green  = int(row['GreenTime'])
    yellow = int(row['YellowTime'])
    red    = int(row['RedTime'])

    total = 0
    curr  = start_dt.date()
    end   = end_dt.date()
    while curr <= end:
        if curr in hols or ((pto_df['ResourceId']==resource_id)
                             & (pto_df['PTODate']==curr)).any():
            total += red
        else:
            total += green
        curr += timedelta(days=1)
    return total

def assess_status(avail, required):
    """
    Determine status:
    - Available: avail >= required
    - At Risk:   avail >= 70% of required
    - Insufficient otherwise
    """
    if avail >= required:
        return 'Available'
    if avail >= required * 0.7:
        return 'At Risk'
    return 'Insufficient'

# Initialize files on startup
ensure_files()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/resources')
def view_resources():
    df = load_resources()
    return render_template('resources.html', resources=df.to_dict('records'))

@app.route('/add_resource', methods=['GET','POST'])
def add_resource():
    if request.method == 'POST':
        df_res = load_resources()
        df_pto = load_pto()
        df_hol = pd.DataFrame({'HolidayDate': load_holidays()})

        new = {
            'ResourceId':   request.form['res_id'].strip(),
            'ResourceName': request.form['res_name'].strip(),
            'GreenTime':    int(request.form.get('green') or DEF_GREEN),
            'YellowTime':   int(request.form.get('yellow') or DEF_YELLOW),
            'RedTime':      int(request.form.get('red') or DEF_RED)
        }
        df_res = pd.concat([df_res, pd.DataFrame([new])], ignore_index=True)
        save_resources(df_res, df_pto, df_hol)
        flash(f"Resource {new['ResourceId']} added.", 'success')
        return redirect(url_for('view_resources'))

    return render_template('add_resource.html')

@app.route('/holidays')
def view_holidays():
    hols = load_holidays()
    return render_template('holidays.html', holidays=hols)

@app.route('/pto')
def view_pto():
    df = load_pto()
    return render_template('pto.html', pto=df.to_dict('records'))

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

    return render_template('project_form.html')

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
        'Status':           status
    }

    df_wi = pd.concat([df_wi, pd.DataFrame([new])], ignore_index=True)
    save_workitems(df_wi)
    flash(f"Assigned {rid} to {name}.", 'success')
    return redirect(url_for('view_workitems'))

@app.route('/workitems')
def view_workitems():
    df = load_workitems()
    df['ProjStart']      = df['ProjStart'].dt.strftime('%Y-%m-%d %H:%M')
    df['ProjEnd']        = df['ProjEnd'].dt.strftime('%Y-%m-%d %H:%M')
    df['AssignDatetime'] = df['AssignDatetime'].dt.strftime('%Y-%m-%d %H:%M')
    return render_template('workitems.html', workitems=df.to_dict('records'))

if __name__ == '__main__':
    app.run(debug=True)
