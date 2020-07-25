from flask import Flask, redirect, url_for, render_template, request, send_from_directory, make_response
import os
import pandas as pd
import numpy as np
import csv
import pdfkit
import datetime as dtm
from datetime import date,datetime



app = Flask(__name__) 
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

APP_ROOT = os.path.dirname(os.path.abspath(__file__))

@app.route("/",methods=["GET", "POST"])
def home():
    return render_template("home.html")
@app.route("/form-sidang",methods=["GET", "POST"])
def index():
    today = date.today()
    dead_rev = today + dtm.timedelta(days=15)
    today = "{:%d-%b-%Y}".format(today)
    dead_rev = "{:%d-%b-%Y}".format(dead_rev)
    now = datetime.now()
    current_time = now.strftime("%H:%M")
    cetak = "0"
    if request.method=="POST":
        try:
            cari = request.form['cari']
            if (cari=="0"):
                return render_template("index.html",  cetak=cetak, message="normal",date=today,dead_rev=dead_rev, current_time=current_time)
        except:
            cari = None
            

        try:
            hitung = request.form['hitung']
        except:
            hitung ="1"

        if (hitung =="1"):
            # return request.form['DPb13']
            NIM = request.form['NIM']
            MHS = request.form['MHS']
            JTA = request.form['JTA']
            Pb1 = request.form['Pb1']
            Pb2 = request.form['Pb2']
            Pg1 = request.form['Pg1']
            Pg2 = request.form['Pg2']

            DPb11 = request.form['DPb11']
            DPb12 = request.form['DPb12']
            DPb13 = request.form['DPb13']
            DPb21 = request.form['DPb21']
            DPb22 = request.form['DPb22']
            DPb23 = request.form['DPb23']

            DPg11 = request.form['DPg11']
            DPg12 = request.form['DPg12']
            DPg13 = request.form['DPg13']
            DPg21 = request.form['DPg21']
            DPg22 = request.form['DPg22']
            DPg23 = request.form['DPg23']
            

            KL = request.form['KL']
            IA = request.form['IA']
            RVS = request.form['RVS']
            time = request.form['current_time']
            ruangan = request.form['ruangan']

            if(current_time!=time and time!=""):
                current_time=time

            try:
                try:
                    cetak = request.form['cetak']
                except:
                    cetak = "0"

                LPb = hitungPembimbing(DPb11,DPb12,DPb13,DPb21,DPb22,DPb23)
                LPg = hitungPenguji(DPg11,DPg12,DPg13,DPg21,DPg22,DPg23)
                LNP = hitungNilaiTotal(LPb,LPg)
                LNA = hitungNilaiAkhir(LNP)
                INA =round( (0.35*LNA[0]) + (0.3*LNA[1]) + (0.35*LNA[2]),2) 
                LIA = indexing(INA)
                html = render_template("index.html",
                                    DPb11=DPb11, DPb12=DPb12, DPb13=DPb13,
                                    DPb21=DPb21, DPb22=DPb22, DPb23=DPb23,
                                    LPb1=LPb[0] , LPb2=LPb[1] , LPb3=LPb[2] ,
                                    DPg11=DPg11, DPg12=DPg12, DPg13=DPg13,
                                    DPg21=DPg21, DPg22=DPg22, DPg23=DPg23,
                                    LPg1=LPg[0] , LPg2=LPg[1] , LPg3=LPg[2] ,
                                    LNP1=LNP[0] , LNP2=LNP[1] , LNP3=LNP[2] ,
                                    LNA1=LNA[0] , LNA2=LNA[1] , LNA3=LNA[2] ,
                                    LIA=LIA, INA=INA, NIM=NIM,
                                    MHS=MHS, JTA=JTA, Pb1=Pb1,
                                    KL=KL,IA=IA,
                                    Pb2=Pb2, Pg1=Pg1, Pg2=Pg2, cetak=cetak,
                                    RVS=RVS,  message="success" ,date=today,
                                    dead_rev=dead_rev, current_time=current_time, ruangan=ruangan)
                # return cetak
                if (cetak=="1"):
                    filename_pdf = "Nilai_"+MHS+".pdf"
                    css = ["static/css/bootstrap.min.css","static/style.css"]
                    config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                    pdf = pdfkit.from_string(html, False,configuration=config, css=css)
                    # pdf = pdfkit.from_string(html, False, css=css)
                    response = make_response(pdf)
                    response.headers["Content-Type"] = "application/pdf"
                    response.headers["Content-Disposition"] = "inline; filename=NilaiSidang.pdf"
                    return response
                else:
                    return html
            except:
                cetak="0"
                return render_template("index.html", cetak=cetak, message="error",date=today,dead_rev=dead_rev, current_time=current_time)

    return render_template("index.html",  cetak=cetak, message="normal",date=today,dead_rev=dead_rev, current_time=current_time)

ind_to_val = {"A":4, "AB":3.5, "B":3, "BC":2.5, "C":2, "D":1, "E":0}
val_to_ind = {4:"A", 3.5:"AB", 3:"B", 2.5:"BC", 2:"C", 1:"D", 0:"E"}

def hitungPembimbing(DPb11,DPb12,DPb13,DPb21,DPb22,DPb23):
    if(DPb11 != "" and (DPb21 != "" or DPb22 != "" or DPb23 != "")):
        # CLO 1
        LPb1 = [ind_to_val[DPb11]]
        LPb1.append(ind_to_val[DPb21])
        LPb1 = np.average(LPb1)
        # CLO 2
        LPb2 = [ind_to_val[DPb12]]
        LPb2.append(ind_to_val[DPb22])
        LPb2 = np.average(LPb2)
        # CLO 3
        LPb3 = [ind_to_val[DPb13]]
        LPb3.append(ind_to_val[DPb23])
        LPb3 = np.average(LPb3)
        LPb=[LPb1,LPb2,LPb3]
        return LPb

    else:
        # CLO 1
        DPb11 = ind_to_val[DPb11]
        LPb1 = DPb11
        # CLO 2
        DPb12 = ind_to_val[DPb12]
        LPb2 = DPb12
        # CLO 3
        DPb13 = ind_to_val[DPb13]
        LPb3 = DPb13
        LPb=[LPb1,LPb2,LPb3]
        return LPb

def hitungPenguji(DPg11,DPg12,DPg13,DPg21,DPg22,DPg23):
    # CLO 1
    LPg1 = [ind_to_val[DPg11]]
    LPg1.append(ind_to_val[DPg21])
    LPg1 = np.average(LPg1)
    # CLO 2
    LPg2 = [ind_to_val[DPg12]]
    LPg2.append(ind_to_val[DPg22])
    LPg2 = np.average(LPg2)
    # CLO 3
    LPg3 = [ind_to_val[DPg13]]
    LPg3.append(ind_to_val[DPg23])
    LPg3 = np.average(LPg3)

    LPg=[LPg1,LPg2,LPg3]
    return LPg

def hitungNilaiTotal(LPb,LPg):   
    # Nilai Hasil Perhituingan
    # CLO1
    LNP1 = (0.6*LPb[0]) + (0.4* LPg[0])
    # CLO2
    LNP2 = (0.6*LPb[1]) + (0.4* LPg[1])
    # CLO3
    LNP3 = (0.6*LPb[2]) + (0.4* LPg[2])
    LNP=[round(LNP1,2),round(LNP2,2),round(LNP3,2)]
    return LNP

def hitungNilaiAkhir(LNP):
    # Nilai Akhir 
    # CLO1
    LNA1 = LNP[0]
    # CLO2
    LNA2 = LNP[1]
    # CLO3
    LNA3 = LNP[2]
    LNA=[LNA1,LNA2,LNA3]
    return LNA

def indexing(nilai_akhir):
    if nilai_akhir > 3.5: LIA = "A"
    elif nilai_akhir > 3.25: LIA = "AB"
    elif nilai_akhir > 2.75: LIA = "B"
    elif nilai_akhir > 2.25: LIA = "BC"
    elif nilai_akhir > 1.75: LIA = "C"
    elif nilai_akhir > 1: LIA = "D"
    else: LIA = "E"
    return LIA



@app.after_request
def add_header(r):
    """
    Add headers to both force latest IE rendering engine or Chrome Frame,
    and also to cache the rendered page for 10 minutes.
    """
    r.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    r.headers["Pragma"] = "no-cache"
    r.headers["Expires"] = "0"
    r.headers['Cache-Control'] = 'public, max-age=0'
    return r

if __name__ == "__main__":
    app.run(debug=True)