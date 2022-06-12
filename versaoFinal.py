
import face_recognition as fr 
import numpy as np
from PIL import Image  
import cv2
from datetime import datetime
import re
import win32com.client as win32
from struct import pack
from tkinter import *
from matplotlib.pyplot import text
from numpy import pad



# cor
co0 = "#f0f3f5"  # Preta / black
co1 = "#feffff"  # branca / white
co2 = "#3fb5a3"  # verde / green
co3 = "#38576b"  # valor / value
co4 = "#403d3d"   # letra / letters


# janela
janela = Tk()
janela.title("Cadastro do Professor")
janela.geometry('310x300')
janela.configure(background=co1)
janela.resizable(width=FALSE, height=FALSE)

#pegar informações
lista_info = []
def cadastrar_emailturma():
    turmaChamada = e_turmaChamada.get()
    emailPrestador = e_emailPrestador.get()
    codigo = len(lista_info)+1
    lista_info.append(turmaChamada)
    lista_info.append(emailPrestador)
    janela.destroy()

# divisão janela
frame_cima = Frame(janela, width=310, height=50, bg=co1, relief='flat')
frame_cima.grid(row=0, column=0, pady=1, padx=0, sticky=NSEW)

frame_baixo = Frame(janela, width=310, height=350, bg=co1, relief='flat')
frame_baixo.grid(row=1, column=0, pady=1, padx=0, sticky=NSEW)

# conf frame_cima
l_cadastro = Label(frame_cima, text='Área do professor', anchor=NE, font=('Ivy 25'), bg=co1, fg=co4)
l_cadastro.place(x=5, y=5)
l_linha = Label(frame_cima, text='', width=275, anchor=NW, font=('Ivy 1'), bg=co2, fg=co4)
l_linha.place(x=10, y=45)

# conf frame_cima
l_email = Label(frame_baixo, text='E-Mail*', anchor=NW, font=('Ivy 15'), bg=co1, fg=co4)
l_email.place(x=10, y=20)
e_emailPrestador = Entry(frame_baixo, width=25, justify='left', font=("", 15), highlightthickness=1, relief='solid')
e_emailPrestador.place(x=14, y=50)

# conf frame_cima
l_turma = Label(frame_baixo, text='Turma*', anchor=NW, font=('Ivy 15'), bg=co1, fg=co4)
l_turma.place(x=10, y=95)
e_turmaChamada = Entry(frame_baixo, width=25, justify='left', font=("", 15), highlightthickness=1, relief='solid')
e_turmaChamada.place(x=14, y=130)

l_cadastrar = Button(frame_baixo, text='Cadastrar', width=39, height=2, font=('Ivy 8 bold'), bg=co2, fg=co1, relief=RAISED, overrelief=RIDGE, command=cadastrar_emailturma)
l_cadastrar.place(x=15, y=200)

janela.mainloop()


turmaChamada = lista_info[0]
emailPrestador = lista_info[1]

print(turmaChamada,emailPrestador)


regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'

def check(eMail):      
    if(re.search(regex,eMail)):  
        #print("Valid Email")  
        return 1
    else:  
        print("Este e-mail não é válido!")


t = check(emailPrestador)



if t ==  1:
    def reconhece_face(url_foto):
        foto = fr.load_image_file(url_foto)
        rostos = fr.face_encodings(foto)
        if (len(rostos)>0) :
            return True, rostos
        
        return False, []


    def get_rostos():
        rostos_conhecidos = []
        nomes_dos_rostos = []
        ra = []

        maria = reconhece_face(r'.\baseFotos\Maria.jpeg')
        if (maria[0]):
            rostos_conhecidos.append(maria[1][0])
            nomes_dos_rostos.append('Maria de Barros Reis')
            ra.append(123456)

        nomes_dos_rostos.append('Desconhecido')
        return     rostos_conhecidos, nomes_dos_rostos, ra

    #---------------------------- Webcam ---------------------------------



    rostos_conhecidos, nome_dos_rostos, ra1 = get_rostos()

    lista_chegadas =[]
    relatorio = []
    preseSalaAula = nome_dos_rostos.copy()
    video_capture = cv2.VideoCapture(0)

    while video_capture.isOpened():
        ret, frame = video_capture.read()

        rgb_frame = frame[:,:,::-1]

        localizacao_dos_rostos = fr.face_locations(rgb_frame)
        rosto_desconhecido = fr.face_encodings(rgb_frame, localizacao_dos_rostos)

        for (top, right, botton, left), face_ecoloding in zip(localizacao_dos_rostos, rosto_desconhecido):
            resultados = fr.compare_faces(rostos_conhecidos, face_ecoloding)
            

            face_distances = fr.face_distance(rostos_conhecidos, face_ecoloding)

            melhor_id = np.argmin(face_distances)
            if resultados[melhor_id]:
                nome = nome_dos_rostos[melhor_id]
                ra = ra1[melhor_id]
                
            else:
                nome = "Desconhecido"
                

            #print(nome)
            # Ao redor do rosto 
            cv2.rectangle(frame, (left,top), (right,botton), (0,0,255),2)

            # #Emabaixo
            cv2.rectangle(frame, (left,botton-35), (right,botton), (0,0,255))
            font = cv2.FONT_HERSHEY_SIMPLEX

            # #texto
            cv2.putText(frame, nome, (left + 6, botton -6 ), font, 1.0, (255,255,255), 1)

            cv2.imshow("WebCam_facerecognition", frame)

            lista = nome[0:]
            
            #print(ra)

            if lista != "Desconhecido" and len(preseSalaAula)>1:

                if lista in preseSalaAula:       

                    
                    rel = datetime.now().strftime("%m-%d-%Y %H:%M")
                    relatorio_now = [lista,rel,ra]
                    relatorio.append(relatorio_now)
                    
                    def gerar_relatorio():
                        print(f"""\nNome do Aluno:\t\tHorario:""")
                        print("-------------------------------------------")
                        for x in range(len(relatorio)):
                            print(f"""{relatorio[x][0][0]}\t\t\t{relatorio[x][1][1]}""")


                    lista_cheg = (datetime.now().strftime("%m-%d-%Y %H:%M"))
                    lista_chegadas.append(lista_cheg)
                    print(lista_chegadas)        
                   

                    now = datetime.now().strftime("%m-%d-%Y %H:%M").replace(':', '-')
                    teste = str(now)
                    x = (nome[0:]+'  '+ teste+' '+'.jpeg')
                    preseSalaAula.remove(lista)
                    cv2.imwrite(x, frame)
        
        if cv2.waitKey(5) == 27:

            dataChamada = datetime.now().strftime("%m-%d-%Y")
            arquivo = open(f'Chamada  {dataChamada} - Turma {turmaChamada}.txt','w')
            arquivo.write(f"\nChamada  {dataChamada} - Turma: {turmaChamada}\n")
            arquivo.write("\n\n")
            arquivo.write("Nome do Aluno:                     Horario:                               RA:\n")
            arquivo.write("-----------------------------------------------------------------------------------------\n")
            for i in range(len(relatorio)):
                arquivo.writelines(str(relatorio[i][0] )+"          " +str(relatorio[i][1] )+"                     " +str(relatorio[i][2] )+'\n')
                
                
            arquivo.close()

            break
        
    video_capture.release()
    cv2.destroyAllWindows()


    #------------------------- Enviar email ------------------------------
    

    if turmaChamada != '' and emailPrestador != '':
        
        # criar a integração com o outlook
        outlook = win32.Dispatch('outlook.application')

        # criar um email
        email = outlook.CreateItem(0)

        DataHoraChamada = datetime.now().strftime("%m-%d-%Y %H:%M")
        

        # configurar as informações do seu e-mail
        email.To = f"{emailPrestador}"
        #email.To = "maria.reis@aluno.faculdadeimpacta.com.br"
        email.Subject = f"E-mail automático - Turma {turmaChamada} - {DataHoraChamada}"
        email.HTMLBody = f"""Bom dia,<br><br>Segue em anexo lista de presença dos alunos da turma {turmaChamada}"""


        dataChamada = datetime.now().strftime("%m-%d-%Y")


        anexo = ('C:/Users/Darly Reis/Desktop/Project/Chamada  '+ dataChamada +' - Turma '+turmaChamada+'.txt')
        print(anexo)
        email.Attachments.Add(anexo)

        email.Send()
        print("Email Enviado")
        


        #-------------------------------- Conexão com o Banco ---------------------------

        import pyodbc


        def retornar_conexao_sql():
            server = #EX - "LAPTOP-SV6" 
            database = "FACULDADEIMPACTA"
            string_conexao = 'Driver={SQL Server Native Client 11.0};Server='+server+';Database='+database+';Trusted_Connection=yes;'
            conexao = pyodbc.connect(string_conexao)
            return conexao.cursor()

        
        for x in range(len(relatorio)):
            id_ra_aluno = str(relatorio[x][2] ) 
            dt_presenca =  str(lista_chegadas[x])
            turmapre = turmaChamada



            cursor = retornar_conexao_sql()
            comando = (f""" set dateformat ymd
            
                            insert into Chamada values({id_ra_aluno}, cast('{dt_presenca}' as datetime) , '{turmapre}' ) """)
            cursor.execute(comando)     
            cursor.commit()
        #cursor.execute(""" SELECT * FROM alunos  """)
        #row = cursor.fetchall()
        #for x in row:
            #print(x)


        print("Salvo dados no banco")