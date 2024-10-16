import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
import pyodbc
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.permanent_session_lifetime = 1800  # Session timeout set to 30 minutes

# Path to PDF directory (modify this based on your actual path)
PDF_DIRECTORY = r'G:\My Drive\DigiDoc\eFax Attachments'

# SQL Server connection string
connection_string = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=10.115.245.191;'  
    'DATABASE=PythonDB;'
    'UID=sa;'
    'PWD=heikinashi'
)

# Ensure PDF storage folder exists
PDF_FOLDER = os.path.join(os.getcwd(), 'pdf_files')
if not os.path.exists(PDF_FOLDER):
    os.makedirs(PDF_FOLDER)

# Database connection
def get_db_connection():
    try:
        return pyodbc.connect(connection_string)
    except Exception as e:
        print(f"Error connecting to SQL Server: {str(e)}")
        return None

# Query PDF data based on the selected date
def query_pdf_data(selected_date):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT ID, Date, Subject, Attachment, Fax_Status, TotalPages 
            FROM dbo.EmailAutomation 
            WHERE Date = ?
        """, (selected_date,))
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

# Parse and convert the date to 'YYYY-MM-DD' format
def format_date(selected_date):
    try:
        formatted_date = datetime.strptime(selected_date, '%Y-%m-%d').date()
        return formatted_date
    except ValueError:
        raise ValueError(f"Incorrect date format for {selected_date}. Expected 'YYYY-MM-DD'.")

# Login route with authentication failure
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        print(f"Attempting login with username: {username}")

        if username == '11693' and password == 'aima@123':
            session['username'] = username
            session.permanent = True  # Set session timeout
            print("Login successful")
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid Credentials')
            print("Login failed")
            return render_template('login.html', error='Invalid credentials. Please try again.')
    return render_template('login.html')

# Dashboard route with default date filter
@app.route('/dashboard', methods=['GET', 'POST'])
def dashboard():
    if 'username' not in session:
        return redirect(url_for('login'))

    conn = None
    cursor = None
    date_filter = request.form.get('date_filter') or datetime.today().strftime('%Y-%m-%d')  # Default current date

    try:
        conn = get_db_connection()
        if conn is None:
            raise Exception("Database connection failed")

        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM dbo.EmailAutomation WHERE Date = ?", (date_filter,))
        total_pdfs = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM dbo.EmailAutomation WHERE Fax_Status = 'Completed' AND Date = ?",
                       (date_filter,))
        completed_files = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM dbo.EmailAutomation WHERE Fax_Status = 'Pending' AND Date = ?",
                       (date_filter,))
        pending_files = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM dbo.EmailAutomation WHERE Fax_Status = 'Exception' AND Date = ?",
                       (date_filter,))
        exception_files = cursor.fetchone()[0]

        cursor.execute("""
            SELECT ID, Date, Subject, Attachment, Fax_Status, TotalPages 
            FROM dbo.EmailAutomation 
            WHERE Date = ?
        """, (date_filter,))
        pdf_data = cursor.fetchall()

        return render_template('dashboard.html',
                               total_pdfs=total_pdfs,
                               completed_files=completed_files,
                               pending_files=pending_files,
                               exception_files=exception_files,
                               pdf_data=pdf_data,
                               selected_date=date_filter)

    except Exception as e:
        flash(f"An error occurred while loading the dashboard: {str(e)}")
        print(f"Dashboard error: {str(e)}")
        return redirect(url_for('login'))
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

# Workbench menu route
@app.route('/workbench', methods=['GET', 'POST'])
def workbench():
    if 'username' not in session:
        return redirect(url_for('login'))

    conn = get_db_connection()
    cursor = conn.cursor()

    # When the form is submitted
    if request.method == 'POST':
        pdf_id = int(request.form['pdf_id'])
        new_status = request.form['status']

        try:
            # Update status in the SQL table
            cursor.execute('UPDATE dbo.EmailAutomation SET Fax_Status = ? WHERE ID = ?', (new_status, pdf_id))
            conn.commit()
            flash('Status updated successfully.')
        except Exception as e:
            flash(f"An error occurred while updating: {str(e)}")
        finally:
            cursor.close()
            conn.close()

    # Fetch the data to display in the Workbench table
    cursor.execute("SELECT ID, Date, Subject, Attachment, Fax_Status, TotalPages FROM dbo.EmailAutomation")
    pdf_data = cursor.fetchall()

    return render_template('workbench.html', pdf_data=pdf_data)

# Serve PDF files
@app.route('/view_pdf/<filename>', methods=['GET'])
def view_pdf(filename):
    file_path = os.path.join(PDF_DIRECTORY, filename)
    if os.path.isfile(file_path):
        return send_file(file_path)
    else:
        flash(f"File {filename} not found")
        return redirect(url_for('dashboard'))

# Update PDF status or delete PDF
@app.route('/update_status', methods=['POST'])
def update_status():
    conn = None
    cursor = None

    try:
        pdf_id = int(request.form['pdf_id'])
        status = request.form['status']
        delete = request.form.get('delete')

        conn = get_db_connection()
        if conn is None:
            raise Exception("Database connection failed")

        cursor = conn.cursor()

        if delete == 'yes':
            cursor.execute('DELETE FROM dbo.EmailAutomation WHERE ID = ?', (pdf_id,))
            flash('PDF deleted successfully.')
        else:
            cursor.execute('UPDATE dbo.EmailAutomation SET Fax_Status = ? WHERE ID = ?', (status, pdf_id))
            flash('Status updated successfully.')

        conn.commit()

    except Exception as e:
        flash(f"An error occurred: {str(e)}")
        print(f"Update Status error: {str(e)}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

    return redirect(url_for('dashboard'))

# Download Excel
@app.route('/download_excel', methods=['GET', 'POST'])
def download_excel():
    selected_date = request.args.get('date_filter') if request.method == 'GET' else request.form.get('date_filter')

    if not selected_date:
        flash("No date provided for filtering the data.")
        return redirect(url_for('dashboard'))

    try:
        formatted_date = format_date(selected_date)
        pdf_data = query_pdf_data(formatted_date)

        if not pdf_data:
            flash(f"No PDF data found for date: {formatted_date}")
            return redirect(url_for('dashboard'))

        pdf_data_list = [{'Date': row[1], 'Subject': row[2],
                          'Attachment': row[3], 'Fax_Status': row[4],
                          'TotalPages': row[5]} for row in pdf_data]

        df = pd.DataFrame(pdf_data_list)

        if df.empty:
            flash("No data to download.")
            return redirect(url_for('dashboard'))

        excel_file = os.path.join(PDF_FOLDER, 'pdf_data.xlsx')
        df.to_excel(excel_file, index=False)
        print(f"Excel file created: {excel_file}")

        return send_file(excel_file, as_attachment=True, download_name='pdf_data.xlsx')

    except Exception as e:
        flash(f"Error generating Excel file: {str(e)}")
        print(f"Excel download error: {str(e)}")
        return redirect(url_for('dashboard'))

# Logout route
@app.route('/logout')
def logout():
    session.pop('username', None)
    flash('You have been logged out.')
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='10.115.245.191', port=5000, debug=True)
