from flask import Flask, render_template, send_file, redirect, request
from flask.helpers import make_response
from docx import Document
from datetime import date

# TODO: FIX THE PROBLEM WITH DOC


# TODO SECOND ENDPOINT


app = Flask(__name__)
global doc
today = date.today()


@app.route('/')
def home():
    return render_template('home.html')


@app.route('/docmaker', methods=['POST'])
def pdf_template():
    global doc
    doctype = request.form.get('doctypes')
    docpath = './templates/documents/'+doctype+'.docx'
    doc = Document(docpath)
    if doctype == 'proofofparticipationinenglishcourses':
        return redirect('/proofofparticipationinenglishcourses')
    elif doctype == 'confirmationofunconditionaladmission':
        return redirect('/confirmationofunconditionaladmission')
    elif doctype == 'extensionletter':
        return redirect('/extensionletter')
    elif doctype == 'latearrival':
        return redirect('/latearrival')
    elif doctype == 'blockedaccount':
        return redirect('/blockedaccount')
    elif doctype == 'onlinesemesterparticipation':
        return redirect('/onlinesemesterparticipation')
    elif doctype == 'presenceletter':
        return redirect('/presenceletter')
    elif doctype == 'proofoflanguagerequirements':
        return redirect('/proofoflanguagerequirements')
    return "Something's wrong"


@app.route('/proofofparticipationinenglishcourses', methods=['GET', 'POST'])
def proofofparticipationinenglishcourses():
    global doc
    if request.method == 'GET':
        return render_template('proofofparticipationinenglishcourses.html')
    else:
        gender = request.form.get('genderattrs')
        name = request.form.get('name') + " " + request.form.get('surname')
        date = today.strftime("%d %B %Y")
        years = request.form.get('yearsattribute')
        studentID = request.form.get('idstudentattribute')
        semester = request.form.get('semesters')
        for p in doc.paragraphs:
            inline = p.runs
            for i in range(len(inline)):
                text = inline[i].text
                if 'nameattribute' in text:
                    text = text.replace('nameattribute', name)
                    inline[i].text = text
                if 'dateattribute' in text:
                    text = text.replace('dateattribute', date)
                    inline[i].text = text
                if 'semesterattribute' in text:
                    text = text.replace('semesterattribute', semester)
                    inline[i].text = text
                if 'yearsattribute' in text:
                    text = text.replace('yearsattribute', years)
                    inline[i].text = text
                if 'genderattribute' in text:
                    print(gender.capitalize())
                    text = text.replace('genderattribute', gender.capitalize())
                    inline[i].text = text
                if 'idstudentattribute' in text:
                    text = text.replace('idstudentattribute', studentID)
                    inline[i].text = text
        doc.save('output.docx')
        return redirect('/download')


# TODO: CAPS LOCK THE COUNTRY!
@app.route('/confirmationofunconditionaladmission', methods=['GET', 'POST'])
def confirmationofunconditionaladmission():
    global doc
    if request.method == 'GET':
        return render_template('confirmationofunconditionaladmission.html')
    else:
        gender = request.form.get('genderattrs')
        name = request.form.get('name') + " " + request.form.get('surname')
        date = today.strftime("%d %B %Y")
        years = request.form.get('yearsattribute')
        studentID = request.form.get('idstudentattribute')
        semester = request.form.get('semesters')
        for p in doc.paragraphs:
            inline = p.runs
            for i in range(len(inline)):
                text = inline[i].text
                if 'nameattribute' in text:
                    text = text.replace('nameattribute', name)
                    inline[i].text = text
                if 'dateattribute' in text:
                    text = text.replace('dateattribute', date)
                    inline[i].text = text
                if 'semesterattribute' in text:
                    text = text.replace('semesterattribute', semester)
                    inline[i].text = text
                if 'yearsattribute' in text:
                    text = text.replace('yearsattribute', years)
                    inline[i].text = text
                if 'genderattribute' in text:
                    print(gender.capitalize())
                    text = text.replace('genderattribute', gender.capitalize())
                    inline[i].text = text
                if 'idstudentattribute' in text:
                    text = text.replace('idstudentattribute', studentID)
                    inline[i].text = text
        doc.save('output.docx')
        return redirect('/download')


@app.route('/extensionletter')
def extensionletter():
    return render_template('extensionletter.html')


@app.route('/latearrival')
def latearrival():
    return render_template('latearrival.html')


# TODO: CHECK the starting date for winter and summer semester to auto-generate it.
@app.route('/blockedaccount')
def blockedaccount():
    return render_template('blockedaccount.html')


@app.route('/onlinesemesterparticipation')
def onlinesemesterparticipation():
    return render_template('onlinesemesterparticipation.html')


@app.route('/presenceletter')
def presenceletter():
    return render_template('presenceletter.html')


@app.route('/proofoflanguagerequirements')
def proofoflanguagerequirements():
    return render_template('proofoflanguagerequirements.html')


@app.route('/download')
def downloadFile():
    return send_file('./output.docx', as_attachment=True)


if __name__ == '__main__':
    app.run(port=5000, debug=True)
