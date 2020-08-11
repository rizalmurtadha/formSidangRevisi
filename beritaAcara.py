from flask import Flask, redirect, url_for, render_template, request, send_from_directory, make_response, session
import os
import pandas as pd
import numpy as np
import csv
import pdfkit
import datetime as dtm
from datetime import date,datetime
import pytz
import joblib
import secrets
import string


app = Flask(__name__)
app.secret_key = "form-sidang-IKN"
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
path_data = os.path.join(APP_ROOT, 'data/')
path_jadwal_sidang = "/".join([path_data,"jadwal_sidang"])
path_output = "/".join([path_data,"output"])
data_admin = "/".join([path_data,"login_admin.p"])
data_recap = "/".join([path_data,"recap.p"])
data_lecturer = "/".join([path_data,"lec_code.p"])
data_schedule = "/".join([path_data,"schedule.p"])
data_rekap_xlsx = "/".join([path_data,"Rekap-Sidang-TA.xlsx"])

# today = date.today()
# dead_rev = today + dtm.timedelta(days=15)
# today = "{:%d-%b-%Y}".format(today)
# dead_rev = "{:%d-%b-%Y}".format(dead_rev)
# # UTC = pytz.utc
# timeZ_Jkt = pytz.timezone('Asia/Jakarta')
# dt_Jkt = datetime.now(timeZ_Jkt)
# # print(dt_Jkt.strftime(' %H:%M:%S %Z %z'))
# current_time = dt_Jkt.strftime('%H:%M')

@app.route("/",methods=["GET", "POST"])
def home():
    if request.method=="POST":
        try:
            cari = request.form['cari']
            # if (cari=="0"):
            #     # button skip
            #     return render_template("index.html",  cetak=cetak, message="normal",date=today,dead_rev=dead_rev, current_time=current_time)
            # else:
            nim=request.form['NIM']
            if(nim==""):
                # return redirect(url_for("home"))
                return render_template("home.html",message="tidak ada")
            passwd_user = request.form["password"]
            try:
                nim = int(nim)
            except:
                return render_template("home.html",message="tidak ada")
            dataMhs = cariMhs(nim,passwd_user)
            if (dataMhs=="tidak ada"):
                return render_template("home.html",message=dataMhs)
            elif (dataMhs=="pwd salah"):
                return render_template("home.html",message=dataMhs)
            elif (dataMhs=="data sudah ada"):
                return render_template("home.html",message=dataMhs)
            else:
                editable="0"
                session["dataMhs0"] = str(dataMhs[0])
                session["dataMhs1"] = dataMhs[1]
                session["dataMhs2"] = dataMhs[2]
                session["dataMhs3"] = dataMhs[3]
                session["dataMhs4"] = dataMhs[4]
                session["dataMhs5"] = dataMhs[5]
                session["dataMhs6"] = dataMhs[6]
                session["dataMhs7"] = dataMhs[7]

                return redirect(url_for("index"))
                # return render_template("index.html", NIM=dataMhs[0], MHS=dataMhs[1],
                #                         JTA=dataMhs[2], pbb1=dataMhs[3], pbb2=dataMhs[4],
                #                         pgj1=dataMhs[5], pgj2=dataMhs[6], ruangan=dataMhs[7],
                #                         cetak=cetak, message="normal",date=today,
                #                         dead_rev=dead_rev, current_time=current_time,editable=editable)
        except:
            cari = None
    try:
        submit = session['submit']
        session.pop('submit',None)
    except:
        submit = "false"
    return render_template("home.html",submit=submit)

@app.route("/login",methods=["GET", "POST"])
def login():
    if "admin" in session:
        return redirect(url_for("admin"))
    else:
        if request.method=="POST":
            try:
                Login = request.form['Login']
            except:
                Login = "0"
            if(Login=="1"):
                login = joblib.load(data_admin)
                login_list = login.username.values.tolist()

                # misal input username di-assign sebagai variable "username"
                username = request.form['username']
                # misal input password di-assign sebagai variable "password"
                password = request.form['password']

                if username not in login_list:
                    # tampilkan kalimat "Username tidak terdaftar"
                    return render_template("login.html",message="error")
                else:
                    user_pwd = login[login.username == username].password[0]
                    if password != user_pwd:
                        # tampilkan kalimat "Password salah"
                        return render_template("login.html",message="pwdSalah")
                    else:
                        # masuk ke halaman admin
                        session["admin"] = username
                        return redirect(url_for("admin"))

        return render_template("login.html")

@app.route("/admin",methods=["GET", "POST"])
def admin():
    if "admin" in session:
        admin = session["admin"]
        if request.method=="POST":
            if request.form['logout']=="1" :
                session.pop('admin')
                return redirect(url_for("home"))
        return render_template("admin.html")
    else:
        return redirect(url_for("login"))
    # return render_template("admin.html")

@app.route("/admin/unduh",methods=["GET", "POST"])
def unduh():
    if "admin" in session:
        admin = session["admin"]
        if request.method=="POST":
            # inputan tanggal awal
            awal = request.form['start']
            awal = awal.replace("-"," ")
            # inputan tanggal akhir
            akhir = request.form['end']
            akhir = akhir.replace("-"," ")
            # memisahkan data tahun, bulan, tanggal
            # urutan data : tahun, bulan, hari/tanggal
            awal=[int(s) for s in awal.split() if s.isdigit()]
            akhir=[int(s) for s in akhir.split() if s.isdigit()]

            # load recap
            recap = joblib.load(data_recap)
            begin = dtm.date(awal[0], awal[1], awal[2]) # input dari date picker kiri (from)
            end = dtm.date(akhir[0], akhir[1], akhir[2]) # input dari date picker kanan (until)
            # saya belum tahu output dari date picker seperti apa, asumsi saya masih bisa diubah ke format datetime
            # filter
            pick_recap = recap[(recap.Tanggal_Ref >= begin) & (recap.Tanggal_Ref <= end)]
            pick_recap = pick_recap.iloc[:,1:]
            # bikin folder baru "output" di dir "data"
            filename = "Hasil-Sidang-{}-{}.xlsx".format(begin,end)
            filename_path = "/".join([path_output,filename])
            # path asli "./data/output/Hasil-Sidang-{}-{}.xlsx".format(begin,end)
            pick_recap.to_excel(filename_path, index=None)
            # selanjutnya file excel yang dihasilkan otomatis terdownload oleh user
            return render_template("unduh.html",download="true", tipe="2",filename=filename)
        return render_template("unduh.html")
    else:
        return redirect(url_for("login"))


@app.route("/admin/unggah",methods=["GET", "POST"])
def unggah():
    if "admin" in session:
        if request.method=="POST":
            file = request.files["file"]
            # load schedule (.p)
            schedule = joblib.load(data_schedule)
            # grab upload excel name
            excel_name = file.filename
            destination = "/".join([path_jadwal_sidang,file.filename])
            file.save(destination)
            # load added schedule
            path_excel = "/".join(["data/jadwal_sidang",excel_name])
            add_schedule =pd.read_excel(path_excel)
            add_schedule.fillna(0, inplace=True)
            col_name = add_schedule.columns.tolist()
            # return str(col_name[5])
            # return str(len(col_name))
            col_name_ref =['No.', 'Nama', 'NIM', 'E-mail', 'KK', 'Pembimbing 1',
                        'Pembimbing 2', 'Penguji 1', 'Penguji 2','Judul',
                        'Tanggal', 'Pukul', 'Skema sidang', 'Lokasi']
            # verify uploaded file is suitable
            if col_name != col_name_ref:
            # tampilkan tulisan "Format file yang diunggah tidak sesuai dengan template"
                return render_template("unggah.html",message="tidak sesuai")
            else:
                inp_schedule = add_schedule.iloc[:,[1,2,5,6,7,8,9,13]]
                # generate password
                list_passwd = gen_passwd(inp_schedule.shape[0])
                # insert password into dataframe
                inp_schedule["Password"] = list_passwd
                add_schedule["Password"] = list_passwd
                # find difference
                schedule_nim_set = set(schedule.NIM.values.tolist())
                inp_schedule_nim_set = set(inp_schedule.NIM.values.tolist())
                diff_ = list(schedule_nim_set - inp_schedule_nim_set)
                diff_schedule = schedule[schedule.NIM.isin(diff_)]
                # concat
                new_schedule = pd.concat([diff_schedule, inp_schedule], axis=0)
                new_schedule.reset_index(drop=True, inplace=True)
                # save to excel
                excel_pwd_name = file.filename.replace(".xlsx","_pwd.xlsx")
                path_excel_pwd = "/".join(["data/jadwal_sidang/",excel_pwd_name])
                add_schedule.to_excel(path_excel_pwd, index=None)
                # save as pickle
                joblib.dump(new_schedule, data_schedule)
                # path_jadwal = os.path.join(APP_ROOT, 'data/jadwal_sidang')
                return render_template("unggah.html",download="true", filenames=excel_pwd_name, tipe="1")
                # return send_from_directory(path_jadwal,filename=excel_pwd_name, as_attachment=True)
                # return render_template("unggah.html",download="true", filenames=excel_pwd_name)
                # return path_excel_pwd
        return render_template("unggah.html")
    else:
        return redirect(url_for("login"))

@app.route('/admin/download<string:filename><string:tipe>')
def download_data(filename,tipe):
    # filename = "Rekap-Sidang-TA.xlsx"
    if tipe == "1":
        # Tipe 1 = file hasil proses unggah ( file jadwal sidang )
        path = path_jadwal_sidang
    elif tipe == "2":
        # Tipe 2 = file hasil proses Unduh ( file hasil sidang )
        path = path_output
    # path_jadwal = os.path.join(APP_ROOT, path)
    return send_from_directory(path,
                               filename=filename, as_attachment=True)

# function for generating password
def gen_passwd(n):
    list_passwd = []
    for i in range(n):
        alphabet = string.ascii_letters + string.digits
        password = ''.join(secrets.choice(alphabet) for i in range(6))
        list_passwd.append(password)
    return list_passwd


@app.route("/form-sidang",methods=["GET", "POST"])
def index():
    today = date.today()
    dead_rev = today + dtm.timedelta(days=15)
    today = "{:%d-%b-%Y}".format(today)
    dead_rev = "{:%d-%b-%Y}".format(dead_rev)
    # UTC = pytz.utc
    timeZ_Jkt = pytz.timezone('Asia/Jakarta')
    dt_Jkt = datetime.now(timeZ_Jkt)
    # print(dt_Jkt.strftime(' %H:%M:%S %Z %z'))
    current_time = dt_Jkt.strftime('%H:%M')
    cetak = "0"
    editable = "1"
    if request.method=="POST":
        # try:
        #     cari = request.form['cari']
        #     if (cari=="0"):
        #         return render_template("index.html",  cetak=cetak, message="normal",date=today,dead_rev=dead_rev, current_time=current_time)
        #     else:
        #         nim=request.form['NIM']
        #         if(nim==""):
        #             # return redirect(url_for("home"))
        #             return render_template("home.html",message="tidak ada")
        #         passwd_user = request.form["password"]
        #         try:
        #             nim = int(nim)
        #         except:
        #             return render_template("home.html",message="tidak ada")
        #         dataMhs = cariMhs(nim,passwd_user)
        #         if (dataMhs=="tidak ada"):
        #             return render_template("home.html",message=dataMhs)
        #         elif (dataMhs=="pwd salah"):
        #             return render_template("home.html",message=dataMhs)
        #         elif (dataMhs=="data sudah ada"):
        #             return render_template("home.html",message=dataMhs)
        #         else:
        #             editable="0"
        #             session["user"] = nim
        #             return render_template("index.html", NIM=dataMhs[0], MHS=dataMhs[1],
        #                                     JTA=dataMhs[2], pbb1=dataMhs[3], pbb2=dataMhs[4],
        #                                     pgj1=dataMhs[5], pgj2=dataMhs[6], ruangan=dataMhs[7],
        #                                     cetak=cetak, message="normal",date=today,
        #                                     dead_rev=dead_rev, current_time=current_time,editable=editable)
        # except:
        #     cari = None

        try:
            hitung = request.form['hitung']
        except:
            hitung ="1"

        if (hitung =="1"):
            # return request.form['DPb13']
            NIM = request.form['NIM']
            MHS = request.form['MHS']
            JTA = request.form['JTA']
            pbb1 = request.form['pbb1']
            pbb2 = request.form['pbb2']
            pgj1 = request.form['pgj1']
            pgj2 = request.form['pgj2']

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

            editable = request.form['editable']
            KL = request.form['KL']
            # IA = request.form['IA']
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
                                    MHS=MHS, JTA=JTA, pbb1=pbb1,
                                    KL=KL,
                                    pbb2=pbb2, pgj1=pgj1, pgj2=pgj2, cetak=cetak,
                                    RVS=RVS,  message="success" ,date=today,
                                    dead_rev=dead_rev, current_time=current_time, ruangan=ruangan, editable=editable)
                # return cetak
                if (cetak=="1"):
                    # recap
                    recap = joblib.load(data_recap)
                    new_today = dtm.datetime.strptime(today, '%d-%b-%Y')
                    recap = recap.append({"Tanggal_Ref": new_today,"Nama": MHS, "NIM": NIM, "Judul": JTA, "Nilai": INA,
                    "Indeks": LIA, "Status": KL, "Tanggal": today, "Waktu": current_time,
                    "Pembimbing 1": pbb1, "Pembimbing 2": pbb2, "Penguji 1": pgj1, "Penguji 2": pgj2}, ignore_index=True)
                    joblib.dump(recap, data_recap)
                    # recap.to_excel(data_rekap_xlsx, index=None)

                    # print pdf
                    filename_pdf = "Form-Sidang-"+NIM+"-"+MHS+".pdf"
                    headers_filename = "attachment; filename="+filename_pdf
                    css = ["static/css/bootstrap.min.css","static/style.css"]
                    # uncomment config yang dipilih
                    # config for heroku
                    # config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
                    config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                    pdf = pdfkit.from_string(html, False,configuration=config, css=css)
                    # config for local pc
                    # pdf = pdfkit.from_string(html, False, css=css)
                    response = make_response(pdf)
                    response.headers["Content-Type"] = "application/pdf"
                    response.headers["Content-Disposition"] = headers_filename
                    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
                    response.headers["Pragma"] = "no-cache"
                    response.headers["Expires"] = "0"
                    response.headers['Cache-Control'] = 'public, max-age=0'
                    return response
                else:
                    return html
            except:
                cetak="0"
                return render_template("index.html", cetak=cetak, message="error",date=today,dead_rev=dead_rev, current_time=current_time,editable=editable)

    if "dataMhs0" in session:
        editable="0"
        lec_code = joblib.load(data_lecturer)
        lec_code.fillna(0, inplace=True)
        lec_name_list = lec_code.Nama.values.tolist()
        lec_name_list.remove(0)
        print(lec_name_list)
        return render_template("index.html", NIM=session['dataMhs0'], MHS=session['dataMhs1'],
                                            JTA=session['dataMhs2'], pbb1=session['dataMhs3'], pbb2=session['dataMhs4'],
                                            pgj1=session['dataMhs5'], pgj2=session['dataMhs6'], ruangan=session['dataMhs7'],
                                            cetak=cetak, message="normal",date=today, lec_name_list=lec_name_list,
                                            dead_rev=dead_rev, current_time=current_time,editable=editable)
    else:
        return redirect(url_for("home"))
        return render_template("index.html",  cetak=cetak, message="normal",date=today,dead_rev=dead_rev, current_time=current_time,editable=editable)

@app.route("/clear")
def clearSession():
    session.pop('dataMhs0',None)
    session.pop('dataMhs1',None)
    session.pop('dataMhs2',None)
    session.pop('dataMhs3',None)
    session.pop('dataMhs4',None)
    session.pop('dataMhs5',None)
    session.pop('dataMhs6',None)
    session.pop('dataMhs7',None)
    session['submit']='true'
    return redirect(url_for('home'))

# cuma buat ngetest data session
@app.route("/test")
def test():
    # return str(session['dataMhs0'])
    return str(session['dataMhs2'])

def cariMhs(nim,passwd_user):
    # [lec_code, schedule] = joblib.load(data_mahasiswa)
    lec_code = joblib.load(data_lecturer)
    schedule = joblib.load(data_schedule)
    recap = joblib.load(data_recap)
    nim_list = schedule.NIM.values.tolist()
    recap_nim = recap.NIM.values.tolist()
    # print(nim in nim_list)

    # misal value NIM yang diisikan di-assign sebagai “nim”
    # condition 1
    if nim not in nim_list:
        # muncul pop up dengan tulisan “NIM Mahasiswa tidak ditemukan di jadwal sidang” dan halaman tidak berubah
        return "tidak ada"
    # condition 2
    else:
        # cek sudah pernah di inputkan atau belum
        if str(nim) in recap_nim:
        # muncul pop up dengan tulisan “Nilai sidang TA Mahasiswa sudah dimasukkan ke database” dan halaman tidak berubah
            return "data sudah ada"
        else:
            # filter data
            sel_data = schedule[schedule.NIM == nim]
            # verifikasi password
            passwd_ref = sel_data["Password"].values[0]
            if passwd_user != passwd_ref:
                # muncul pop up dengan tulisan “Password salah” dan halaman tidak berubah
                return "pwd salah"
            # grab nama
            else:
                nama = sel_data["Nama"].values[0]
                # grab judul
                judul = sel_data["Judul"].values[0]
                # list kode dosen
                lec_code_list = lec_code.Kode.values.tolist()
                lec_name_list = lec_code.Nama.values.tolist()
                # lec_name_list.remove("nan")
                # grab nama pembimbing 1
                pbb1 = sel_data["Pembimbing 1"].values[0]
                if pbb1 in lec_code_list:
                    pbb1 = lec_code[lec_code.Kode == pbb1].Nama.values[0]
                # grab nama pembimbing 2
                pbb2 = sel_data["Pembimbing 2"].values[0]
                if pbb2 != 0:
                    if pbb2 in lec_code_list:
                        pbb2 = lec_code[lec_code.Kode == pbb2].Nama.values[0]
                else:
                    pbb2 = ""
                # grab nama penguji 1
                pgj1 = sel_data["Penguji 1"].values[0]
                if pgj1 in lec_code_list:
                    pgj1 = lec_code[lec_code.Kode == pgj1].Nama.values[0]
                # grab nama penguji 2
                pgj2 = sel_data["Penguji 2"].values[0]
                if pgj2 in lec_code_list:
                    pgj2 = lec_code[lec_code.Kode == pgj2].Nama.values[0]
                # grab lokasi sidang
                lokasi = sel_data["Lokasi"].values[0]
                # kirim variable ke halaman selanjutnya
                return [nim, nama, judul, pbb1, pbb2, pgj1, pgj2, lokasi]


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

@app.route('/admin/<string:filename>')
def download_files(filename):
    # filename = "Rekap-Sidang-TA.xlsx"
    return send_from_directory(path_data,
                               filename=filename, as_attachment=True)

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
