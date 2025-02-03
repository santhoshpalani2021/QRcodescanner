from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import time
import os

app = Flask(__name__)

# Load the remote log (or create an empty one)
def load_remote_log():
    try:
        df = pd.read_excel("remote_log.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Faculty Name","Remote ID", "Date of Issue", "Time of Issue", "Return Status", "Date of Return", "Time of Return"])
        df.to_excel("remote_log.xlsx", index=False)
    return df

# Home route - shows remote list
@app.route('/')
def index():
    df = load_remote_log()
    remotes = df.iloc[::-1].to_dict(orient='records')  # Convert DataFrame to a list of dicts
    return render_template('index.html', remotes=remotes)

# Route to handle Remote Collection (Issuing)
@app.route('/remote_collection', methods=['GET', 'POST'])
def remote_collection():
    # Load the remote log and prepare remotes data for display
    df = load_remote_log()
    remotes = df.tail(10).iloc[::-1].to_dict(orient='records')  # Convert DataFrame to a list of dicts

    if request.method == 'POST':
        faculty_name = request.form['faculty_name']
        remote_id = request.form['remote_id']
        issue_time = time.strftime("%H:%M:%S")
        issue_date = time.strftime("%Y-%m-%d")

        # Log the issue
        new_entry = {
            "Faculty Name": faculty_name,
            "Remote ID": remote_id,
            "Date of Issue": issue_date,
            "Time of Issue": issue_time,
            "Return Status": "Not Returned",
            "Date of Return": "",
            "Time of Return": ""
        }
        df = df.append(new_entry, ignore_index=True)
        df.to_excel("remote_log.xlsx", index=False)

        return redirect(url_for('remote_collection'))  # Redirect to the remote collection page to view updated list

    return render_template('remote_collection.html', remotes=remotes)  # Pass remotes data to template
# Route to handle Remote Return

@app.route('/remote_return', methods=['GET', 'POST'])
def remote_return():
    # Load the remote log and prepare remotes data for display
    df = load_remote_log()

    # Only include remotes that are "Not Returned"
    remotes = df[df["Return Status"] == "Not Returned"].tail(10).iloc[::-1].to_dict(orient='records')  # Get last 10 Not Returned

    if request.method == 'POST':
        remote_id = request.form['remote_id']

        # Log the return
        index = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")].index
        if not index.empty:
            return_time = time.strftime("%H:%M:%S")
            return_date = time.strftime("%Y-%m-%d")
            df.loc[index, "Return Status"] = "Returned"
            df.loc[index, "Date of Return"] = return_date
            df.loc[index, "Time of Return"] = return_time
            df.to_excel("remote_log.xlsx", index=False)

        return redirect(url_for('remote_return'))  # Redirect to the remote return page to view updated list

    return render_template('remote_return.html', remotes=remotes)  # Pass remotes data to the template
# Route to download the Excel file
@app.route('/download_excel')
def download_excel():
    # Ensure that the file exists
    if os.path.exists("remote_log.xlsx"):
        return send_file("remote_log.xlsx", as_attachment=True)
    else:
        return "File not found", 404  # Return a 404 if file doesn't exist

# Main entry point to run the app
if __name__ == '__main__':
    app.run(debug=True)