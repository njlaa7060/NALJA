from flask import Flask, render_template, request, redirect, url_for, session,flash, Blueprint, render_template, request, send_file, send_from_directory
from werkzeug.utils import secure_filename
from io import BytesIO
import os
import re
import json, time

# Library Convert doc, ppt, xls to PDF
import comtypes.client
import win32com.client as win32

# get PDF contents using PyMuPdf library called as fitz
# more info at https://pymupdf.readthedocs.io/en/latest/tutorial.html
import fitz

# import translation API
# more info at https://stackabuse.com/text-translation-with-google-translate-api-in-python/
import googletrans
from googletrans import Translator

# import url data from youtube
import requests
from bs4 import BeautifulSoup


from .dbase import db

views = Blueprint('views', __name__)

# connect to database
con = db.connect()

local = "Yes"

BASE_DIR = '/home/CES/mysite/website/'
if local == "Yes":
    BASE_DIR = "website/" #'/home/CES/mysite/website/'

global langs

# define a translator
translator = Translator()


# Set Site Language
@views.route('/setLang/<string:lang>', methods=['GET', 'POST'])
def setLang(lang):
    session['lang'] = lang; # ar or en
    return redirect(url_for('views.home'))

# ======================
# Convert file to PDF 
# ======================
def ConvertToPDF(srcfile):
    fileprts = srcfile.split(".")
    fname = fileprts[0]
    newfile = os.path.abspath(fname + ".pdf") #new pdf file name
    fileext = fileprts[1]
    objtype = ""
    comtypes.CoInitialize()
    # Convert Microsoft Word
    if fileext == "docx":
        obj = win32.Dispatch("Word.Application")
        doc = obj.Documents.Open(os.path.abspath(srcfile))
        doc.SaveAs(newfile, 17)
        doc.Close()
        obj.Quit()
    # Convert Microsoft Excel
    if  fileext == "xlsx":
        excel = win32.Dispatch("Excel.Application")        
        sheets = excel.Workbooks.Open(os.path.abspath(srcfile))
        work_sheets = sheets.Worksheets[0]
        work_sheets.ExportAsFixedFormat(0, newfile)
    # Open PowerPoint document and convert to PDF
    if  fileext == "pptx":
        powerpoint = win32.Dispatch('Powerpoint.Application')
        presentation = powerpoint.Presentations.Open(os.path.abspath(srcfile))
        presentation.SaveAs(newfile , 32)
        presentation.Close()
        powerpoint.Quit()
    return newfile

# ======================
# read PDF file
# ======================
def readPDF(filename):
    contents = []
    pagenum = 0
    doc = fitz.open(filename)
    for page in doc:
        pagenum += 1
        # do something with 'page'
        contents.append(page.get_text("text"))
    return [contents, pagenum]

def getLang():
    if not "lang" in session:
        session['lang'] = "en"
    f = open(BASE_DIR + "static/langs/" + session['lang'] + ".json", "r", encoding="utf-8")
    print(BASE_DIR + "static/langs/" + session['lang'] + ".json")
    langs = json.loads(f.read())
    return langs

# ====================== Start Site func ============================================
@views.route('/')
def index():
    if not "lang" in session:
        session['lang'] = "en"
    return redirect(url_for('views.home'))

@views.route('/home')
def home():
    langs = getLang()
    # Check if user is loggedin
    if 'loggedin' in session:
        qry = "SELECT * FROM tr_users WHERE username='" + session['username'] + "'"
        users = db.selectQry(con, qry)
        # User is loggedin show them the home page
        return render_template('site/home.html', sessioninfo=langs)
    # User is not loggedin redirect to login page
    return render_template('site/home.html', sessioninfo=langs)

@views.route('/aboutus')
def aboutus():
    langs = getLang()
    return render_template('site/aboutus.html', sessioninfo=langs)

@views.route('/contactus')
def contactus():
    langs = getLang()
    return render_template('site/contactus.html', sessioninfo=langs)


# http://localhost:5000/register
# This will be the registration page, we need to use both GET and POST requests
@views.route('/register', methods=['GET', 'POST'])
def register():
    langs = getLang()
    # Check if "username", "password" and "email" POST requests exist (user submitted form)
    if request.method == 'POST' and 'username' in request.form and 'password' in request.form and 'email' in request.form:
        # Create variables for easy access
        username = request.form['username']
        password = request.form['password']
        fullName = request.form['fullname']
        email = request.form['email']

        # Check if account exists using MySQL
        # cursor.execute('SELECT * FROM accounts WHERE username = %s', (username))
        qry = "SELECT * FROM tr_users WHERE username='" + username + "'"
        account = db.selectQry(con, qry)

        # If account exists show error and validation checks
        if account[1] == 1:
            flash("Account already exists!", "danger")
        elif not re.match(r'[^@]+@[^@]+\.[^@]+', email):
            flash("Invalid email address!", "danger")
        elif not re.match(r'[A-Za-z0-9]+', username):
            flash("Username must contain only characters and numbers!", "danger")
        elif not username or not password or not email:
            flash("empty username/password!", "danger")
        else:
            # Account doesnt exists and the form data is valid, now insert new account into accounts table
            qry = "INSERT INTO tr_users (username, password, fullname, email, usertypeid) "
            qry += " VALUES ('" + username + "','" + password + "','" + fullName + "','" + email + "', '2')"
            db.crudQry(con, qry)
            flash("You have successfully registered!", "success")
            return redirect(url_for('views.login'))

    elif request.method == 'POST':
        # Form is empty... (no POST data)
        flash("Please fill out the form!", "danger")
    # Show registration form with message (if any)
    return render_template('auth/register.html', sessioninfo=langs)


# http://localhost:5000/ - this will be the login page, we need to use both GET and POST requests
@views.route('/login', methods=['GET', 'POST'])
def login():
    langs = getLang()
    # Check if "username" and "password" POST requests exist (user submitted form)
    if request.method == 'POST' and 'username' in request.form and 'password' in request.form:
        # Create variables for easy access
        username = request.form['username']
        password = request.form['password']
        qry = "SELECT * FROM tr_users WHERE username='" + username + "' AND password='" + password + "'"
        users = db.selectQry(con, qry)
        for user in users[0]:
            account = user
        if users[1] > 0:
            # Create session data, we can access this data in other routes
            session['loggedin'] = True
            session['userid'] = account['userid']
            session['usertypeid'] = account['usertypeid']
            session['username'] = account['username']
            return redirect(url_for('views.home'))
        else:
            # Account doesnt exist or username/password incorrect
            return render_template('auth/login.html', sessioninfo=langs, msg=langs['loginerrmsg'])
    return render_template('auth/login.html', sessioninfo=langs, msg="")

@views.route('/logout')
def logout():
    langs = getLang()
    session["id"] = ""
    session["loggedin"] = False
    session["usertypeid"] = ""
    session["username"] = ""
    session.clear()
    return render_template('auth/login.html', sessioninfo=langs)

@views.route('/profile/<string:mode>', methods=['GET', 'POST'])
def profile(mode="view"):
    langs = getLang()
    # update the user profile
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        fullname = request.form["fullname"]
        email = request.form["email"]
        # Update Query
        qry = "UPDATE tr_users SET "
        qry += "username='" + username + "', password='" + password + "', fullname='" + fullname + "', email='" + email + "' "
        qry += "WHERE userid='" + str(session["userid"]) + "'"
        db.crudQry(con, qry)
        return redirect(url_for("views.profile", mode="view"))
    else:
        # Check if user is loggedin
        if 'loggedin' in session:
            qry = "SELECT * FROM tr_users WHERE username='" + session['username'] + "'"
            users = db.selectQry(con, qry)
            for user in users[0]:
                account = user
            # User is loggedin show them the home page
            return render_template('auth/profile.html', sessioninfo=langs, account=account, mode=mode)    
    # User is not loggedin redirect to login page
    return redirect(url_for('views.login'))

@views.route('/dashboard')
def dashboard():
    langs = getLang()
    userid = session['userid']
    if 'loggedin' in session:
        qry = "SELECT * FROM tr_files WHERE userid = '" + str(userid) + "'"
        result = db.selectQry(con, qry)
        filerows = result[0]
        counts = result[1]
        # get the translation data
        qry = "SELECT * from tr_translates WHERE fileid IN (SELECT fileid FROM tr_files WHERE userid = '" + str(userid) + "')"
        result2 = db.selectQry(con, qry)
        transrows = result2[0]
        # get the shared files
        shareqry = "SELECT * from tr_translates t "
        shareqry += " LEFT JOIN tr_files f ON f.fileid = t.fileid "
        shareqry += " WHERE t.sharedwith LIKE '%" + str(userid) + "%'"
        result3 = db.selectQry(con, shareqry)
        resrows = result3[0]
        sharerows = []
        for resrow in resrows:
            if resrow["sharedwith"] is not None:
                sharedids = resrow["sharedwith"].split(",") #2,3,
                for sharedid in sharedids:
                    if str(sharedid) == str(userid):
                        sharerows.append(resrow)

        return render_template('auth/dashboard.html', sessioninfo=langs, filerows=filerows, counts=counts, transrows=transrows, sharerows=sharerows)

    return redirect(url_for('views.login'))

########################################## UPLOADING A File ###################################
# upload folder
UPLOAD_FOLDER = BASE_DIR + 'uploads'
Download_FOLDER = BASE_DIR + 'download'

# allowed files to upload
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'xlsx', 'pptx'}
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@views.route('/upload_file', methods=['GET', 'POST'])
def upload_file():
    uploadtime = time.time()
    userid = session["userid"]
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            return "No file part exists"
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            return "No file selected "
        if file and allowed_file(file.filename):
            # Save file on disk
            fileext = file.filename.rsplit('.', 1)[1].lower()
            filename = secure_filename(file.filename)
            file.save(UPLOAD_FOLDER + "/" + filename)
            file_stats = os.stat(UPLOAD_FOLDER + "/" + filename)
            filesize = file_stats.st_size
            client_filename = file.filename.rsplit('.', 1)[0].lower()
            # Save file as record  in database
            qry = "INSERT INTO tr_files (userid,client_filename,filename,fileext,filesize,createdon) " 
            qry += " VALUES ('" + str(userid) + "','" + str(client_filename) + "','" + str(filename) + "','" + str(fileext) + "', '" + str(filesize) + "', '" + str(uploadtime)+ "')"
            db.crudQry(con, qry)
            return "done"
    return "error"

#######################
# Delete a file
#######################
@views.route("/delFile", methods=['GET', 'POST'])
def delFile():
    if request.method == "POST":
        recid = request.form["recid"]
        # delete file from folder 
        qry = "SELECT * FROM tr_files WHERE fileid = '" + str(recid) + "'"
        rows = db.selectQry(con, qry)
        filename = ""
        for row in rows[0]:
            filename = row["filename"]
        if filename != "":
            filepath = UPLOAD_FOLDER + "/" + filename
            os.remove(filepath) 

        # Delete file from database by fileid
        qry = "DELETE FROM tr_files WHERE fileid = '" + str(recid) + "'"
        db.crudQry(con, qry)
        return "done"
    return "error"

#######################
# Translate file contents
#######################
@views.route('/translateFile/<string:fileid>')
def translateFile(fileid):
    langs = getLang()
    if 'loggedin' in session:
        qry = "SELECT * FROM tr_files WHERE fileid = '" + str(fileid) + "'"
        result = db.selectQry(con, qry)
        filerows = result[0]
        fileinfos = ["", 0]
        for filerow in filerows:
            if filerow["fileext"] == "pdf":
                fileinfos = readPDF(UPLOAD_FOLDER + "/" + filerow["filename"])
            else:
                pdffile = ConvertToPDF(UPLOAD_FOLDER + "/" + filerow["filename"])
                fileinfos = readPDF(pdffile)
        Translates = []
        return render_template('auth/translateFile.html', sessioninfo=langs, filerows=filerows, fileid=fileid, fileinfos=fileinfos, Translates=Translates)
    return redirect(url_for('views.login'))

# ======================
# Translate a text
# ======================
@views.route('/translate', methods=['GET', 'POST'])
def translate():
    if 'loggedin' in session:
        content = request.form['content']
        src = request.form['src'] # from lang
        dest = request.form['dest'] # to lang
        TransPage = translator.translate(content, src=src, dest=dest)
        return TransPage.text
    return redirect(url_for('views.login'))

# =====================
# Find in youtube
# =====================
@views.route("/findInYouTube", methods=['GET', 'POST'])
def findInYouTube():
    if 'loggedin' in session:
        # content = request.form['content']
        # requesting url
        headers = {'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'}

        textToSearch = request.form['content']
        url = 'https://www.youtube.com/results'

        response = requests.get(url, params={'search_query': textToSearch}, headers=headers)
        parts = response.text.split("</script>")
        objparts = parts[33].split(">")
        finparts = objparts[1].split("= ")
        latparts = finparts[1][:-1]
        print(parts[33])
        return latparts
    return redirect(url_for('views.login'))


##########################################
# save translation in database
##########################################
@views.route("/saveTransFile", methods=["GET", "POST"])
def saveTransFile():
    if 'loggedin' in session:
        if request.method == "POST":
            filecontent = request.form["filecontent"].replace("'", "\\'")
            filecontent = re.sub('<[^<]+?>', '', filecontent)
            filecontent = filecontent.replace('\n\n','\n')
            fileid = request.form["fileid"]
            # check on previous saved translated data
            qry = "SELECT * from tr_translates WHERE fileid = '" + str(fileid) + "'"
            foundrows = db.selectQry(con, qry)
            # if no previous data
            if foundrows[1] == 0:
                # insert data
                insertqry = "INSERT INTO tr_translates (fileid, translates) VALUES ('" + str(fileid) + "','" + str(filecontent) + "')"
                db.crudQry(con, insertqry)
                status = "Saved"
            else:
                # update data
                updateqry = "UPDATE tr_translates SET translates = '" + str(filecontent) + "' WHERE fileid='" + str(fileid) + "' "
                db.crudQry(con, updateqry)
                status = "Updated"
            print(status)
            return status
    return redirect(url_for('views.login'))
    
##########################################
# view translation
##########################################
@views.route("/viewTranslate/<string:fileid>", methods=["GET", "POST"])
def viewTranslate(fileid):
    langs = getLang()
    if 'loggedin' in session:
        # check on previous saved translated data
        qry = "SELECT * from tr_translates WHERE fileid = '" + str(fileid) + "'"
        result = db.selectQry(con, qry)
        transrows = result[0]
        return render_template('auth/viewTranslate.html', sessioninfo=langs, transrows=transrows)
    return redirect(url_for('views.login'))

##########################################
# save the file 
##########################################
@views.route("/downloadFile/<string:fileid>", methods=["GET", "POST"])
def downloadFile(fileid):
    qry = "SELECT * from tr_translates WHERE fileid = '" + str(fileid) + "'"
    result = db.selectQry(con, qry)
    transrows = result[0]
    contents = ""
    for transrow in transrows:
        contents = transrow["translates"]
    filename = str(int(time.time())) + ".txt" 
    url = os.path.join(views.root_path, "download\\" + filename)
    saveFile(filename, contents)
    return send_file(url, as_attachment=True )

############################
# Save new  file as text
############################
def saveFile(filename, contents):
    f = open(Download_FOLDER + "/" + filename, "w", encoding="utf-8")
    f.write(contents.replace("\n\n","\n"))
    f.close()

############################
# get current system users 
############################
@views.route("/getUsers", methods=["GET", "POST"])
def getUsers():
    qry = "SELECT * FROM tr_users WHERE userid != '" + str(session["userid"]) + "'"
    result = db.selectQry(con, qry)
    users = result[0]

    ops = 'Please select user: <select name="seluserid" id="seluserid">'
    ops += '<option value="all">All</option>'
    for user in users:
        ops += '<option value="' + str(user["userid"]) + '">' + str(user["username"]) + '</option>'
    ops += '</select>'
    return ops

############################
# save shared users
############################
@views.route("/saveShare", methods=["GET", "POST"])
def saveShare():
    fileid = request.form["fileid"]
    shareid = request.form["shareid"] # userid
    qry = "SELECT * from tr_translates WHERE fileid = '" + str(fileid) + "'"
    result = db.selectQry(con, qry)
    transrows = result[0]
    newshare = []
    newshare.append(shareid)
    oldshare = ""
    for transrow in transrows:
        if transrow["sharedwith"] is not None:
            oldshare = transrow["sharedwith"].split(",")
    if oldshare is not None:
        for old in oldshare:
            if old not in newshare:
                newshare.append(old)

    qry = "UPDATE tr_translates SET sharedwith = '" + ",".join(newshare) + "' WHERE fileid = '" + str(fileid) + "'"
    db.crudQry(con, qry)
    return "done"



#========================================= Admin area ====================================================
@views.route('/viewusers')
def viewusers():
    langs = getLang()
    # Check if user is loggedin
    if 'loggedin' in session:
        qry = "SELECT * FROM tr_users u"
        qry += " LEFT JOiN tr_usertypes g ON g.usertypeid = u.usertypeid"
        users = db.selectQry(con, qry)
        # User is loggedin show them the home page
        return render_template('admin/viewusers.html', sessioninfo=langs, users=users)
    # User is not loggedin redirect to login page
    return redirect(url_for('views.login'))

@views.route('/newuser', methods=['GET', 'POST'])
def newuser():
    langs = getLang()
    # Check if user is loggedin
    if 'loggedin' in session:
        if request.method == 'POST':
            username = request.form['username']
            password = request.form['password']
            email = request.form['email']
            fullName = request.form['fullname']
            usertypeid = request.form['usertypeid']
            qry = "INSERT INTO tr_users (username, password, fullname, email, usertypeid) "
            qry += " VALUES ('"+username+"','"+password+"','"+fullName+"','"+email+"','" + usertypeid + "')"
            db.crudQry(con, qry)
            return redirect(url_for('views.viewusers'))
        else:
            qry2 = "SELECT * FROM tr_usertypes"
            groups = db.selectQry(con, qry2)
            # User is loggedin show them the home page
            return render_template('admin/newuser.html', sessioninfo=langs, groups=groups)
    # User is not loggedin redirect to login page
    return redirect(url_for('views.login'))

@views.route('/edituser/<string:userid>', methods=['GET', 'POST'])
def edituser(userid):
    langs = getLang()
    # Check if user is loggedin
    if 'loggedin' in session:
        if request.method == 'POST':
            username = request.form['username']
            password = request.form['password']
            email = request.form['email']
            fullName = request.form['fullname']
            usertypeid = request.form['usertypeid']
            qry = "UPDATE tr_users SET username='"+username+"', password='"+password+"', "
            qry += "email='"+email+"', fullname='"+fullName+"', usertypeid='"+usertypeid+"' WHERE userid='" + userid + "'"
            db.crudQry(con, qry)
            return redirect(url_for('views.viewusers'))
        else:
            qry = "SELECT * FROM tr_users WHERE userid='" + userid + "'"
            users = db.selectQry(con, qry)
            qry2 = "SELECT * FROM tr_usertypes"
            groups = db.selectQry(con, qry2)
            # User is loggedin show them the home page
            return render_template('admin/edituser.html', sessioninfo=langs, groups=groups, users=users)
    # User is not loggedin redirect to login page
    return redirect(url_for('views.login'))


@views.route('/deluser/<string:userid>', methods=['GET', 'POST'])
def deluser(userid):
    langs = getLang()
    # Check if user is loggedin
    if 'loggedin' in session:
        if request.method == 'POST':
            qry = "DELETE FROM tr_users WHERE userid='" + userid + "'"
            db.crudQry(con, qry)
            return redirect(url_for('views.viewusers'))
        else:
            qry = "SELECT * FROM tr_users WHERE userid='" + userid + "'"
            users = db.selectQry(con, qry)
            qry2 = "SELECT * FROM tr_usertypes"
            groups = db.selectQry(con, qry2)
            # User is loggedin show them the home page
            return render_template('admin/deluser.html', sessioninfo=langs, groups=groups, users=users)
    # User is not loggedin redirect to login page
    return redirect(url_for('views.login'))




