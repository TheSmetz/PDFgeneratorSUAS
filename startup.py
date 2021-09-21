from flask import Flask, render_template, send_file, redirect, request
from flask.helpers import make_response
from docx import Document

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('home.html')


# BE CARE ABOUT DOC AND NOT DOCX
@app.route('/docmaker', methods=['POST'])
def pdf_template():
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
    # name = request.form.get('name')
    # date = request.form.get('date')
    # years = request.form.get('years')
    # semester = request.form.get('semester')
    # for p in doc.paragraphs:
    #     inline = p.runs
    #     for i in range(len(inline)):
    #         text = inline[i].text
    #         if 'nameinput' in text:
    #             text=text.replace('nameinput',name)
    #             inline[i].text = text
    #         if 'dateinput' in text:
    #             text=text.replace('dateinput',date)
    #             inline[i].text = text
    #         if 'semesterinput' in text:
    #             text=text.replace('semesterinput',semester)
    #             inline[i].text = text
    #         if 'yearsinput' in text:
    #             text=text.replace('yearsinput',years)
    #             inline[i].text = text
    # doc.save('output.docx')
    # return redirect('/download')


@app.route('/proofofparticipationinenglishcourses')
def proofofparticipationinenglishcourses():
    return render_template('proofofparticipationinenglishcourses.html')
    

# TODO: CAPS LOCK THE COUNTRY!
@app.route('/confirmationofunconditionaladmission')
def confirmationofunconditionaladmission():
    return render_template('confirmationofunconditionaladmission.html')

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
def downloadFile ():
    return send_file('./output.docx', as_attachment=True)

if __name__ == '__main__':
    app.run(port=5000,debug=True) 