import pandas as pd
import requests
import io
import warnings
warnings.filterwarnings("ignore")

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import datetime

mapa = {
'Ene*':'Ene',
'Feb*':'Feb',
'Mar*':'Mar',
'Abr*':'Abr',
'May*':'May',
'Jun*':'Jun',
'Jul*':'Jul',
'Ago*':'Ago',
'Sep*':'Sep',
'Oct*':'Oct',
'Nov*':'Nov',
'Dic*':'Dic'
}

def redeterminacion():
    
    #Extraigo base INDEC
    
    df = pd.read_excel('https://www.indec.gob.ar/ftp/cuadros/economia/series_sipm_dic2015.xls', sheet_name=1)
    
    # Limpio df INDEC

    df.iloc[3] = df.iloc[3].replace(mapa)
    df.iloc[2].fillna(method='ffill', inplace = True)
    df.iloc[2] = df.iloc[3].astype(str).str.cat(df.iloc[2].astype(str), sep=' ')
    df.columns = df.iloc[2]
    df = df[6:]
    df = df.dropna(thresh=df.shape[1]-2)
    df = df.reset_index(drop=True)
    df.rename(columns={'nan Código':'Codigo', 'nan Descripción':'Descripcion'}, inplace=True)
    df.columns.names = ['']
    
    # Calculo variaciones de categorías relevantes

    maq_eq = df[df['Codigo']==29]
    maq_eq_aumento = (maq_eq.iloc[:,-1]/maq_eq.iloc[:,-2]*100)-100
    
    petro = df[df['Codigo']==23]
    petro_aumento = (petro.iloc[:,-1]/petro.iloc[:,-2]*100)-100
    
    cauyplas = df[df['Codigo']==25]
    cauyplas_aumento = (cauyplas.iloc[:,-1]/cauyplas.iloc[:,-2]*100)-100
    
    textil = df[df['Codigo']==18]
    textil_aumento = (textil.iloc[:,-1]/textil.iloc[:,-2]*100)-100
    
    ng = df[df['Codigo']=='NG']
    ng_aumento = (ng.iloc[:,-1]/ng.iloc[:,-2]*100)-100
    
    # Genero df de resultados

    df1 = pd.DataFrame(columns=['Peso',df.columns[-2], df.columns[-1], 'Variacion', 'Total'], index=[
                                                                                                 'IPIM 7.2.1 - Máquinas y Equipos (29)',
                                                                                                 'IPIM 7.2.1 - Productos Refinados de Petróleo (23)',
                                                                                                 'IPIM 7.2.1 - Productos de Caucho y Plástico (25)',
                                                                                                 'IPIM 7.2.1 - Prendas de Materiales y Textiles (18)',
                                                                                                 'Nivel General'])
    
    # Agrego valores al df de resultados
    
    peso = pd.Series([0.2, 0.11, 0.02, 0.01, 0.1])
    df1['Peso'] = peso.values
    anterior = pd.Series([maq_eq.iloc[:,-2].item(), petro.iloc[:,-2].item(), cauyplas.iloc[:,-2].item(), textil.iloc[:,-2].item(), ng.iloc[:,-2].item()])
    df1[df.columns[-2]] = anterior.values
    actual = pd.Series([maq_eq.iloc[:,-1].item(), petro.iloc[:,-1].item(), cauyplas.iloc[:,-1].item(), textil.iloc[:,-1].item(), ng.iloc[:,-1].item()])
    df1[df.columns[-1]] = actual.values
    df1['Variacion'] = (df1[df.columns[-1]]/df1[df.columns[-2]])*100-100
    df1['Total'] = df1['Variacion'] * df1['Peso']
    
    # Extraigo escalas salariales

    url = 'https://sindicatojardineros.org/convenios/espacios-verdes.html'
    res = requests.get(url)
    df_list = pd.read_html(res.content, match='Oficial de espacios verdes')
    df2 = df_list[0]

    # Guardo categoría salarial relevante y convierto a numérico los valores

    df2 = df2[df2['Categoría']=='Oficial de espacios verdes'].reset_index(drop=True)
    df2.iloc[:,1:] = pd.Series(df2.iloc[:,1:].values.flatten()).str.replace('.', ' ')
    df2.iloc[:,1:] = pd.Series(df2.iloc[:,1:].values.flatten()).str.replace(',', '.')
    df2.iloc[:,1:] = pd.Series(df2.iloc[:,1:].values.flatten()).str.replace(' ', '')
    
    cols = df2.columns
    df2[cols[1:]] = df2[cols[1:]].apply(pd.to_numeric, errors='coerce')
    
    # Concateno df INDEC y df Salario
    
    try:
        match = df2.columns[df2.columns.isin(df1.columns)]
        column_index = df2.columns.get_loc(match.item())
        df2 = df2.iloc[:, column_index - 1:column_index + 1]
        df2.rename(columns={df2.columns[0]:df1.columns[1]}, inplace=True)
        df2['Variacion'] = (df2[df2.columns[-1]]/df2[df2.columns[-2]])*100-100
        df2['Peso'] = 0.55
        df2['Total'] = df2['Variacion'] * df2['Peso']
        df2 = df2[df1.columns]
        df2.rename(index={0:'Básico Convenio Colectivo N° 653/12 - Categoría "Oficial de Espacios Verdes"'}, inplace=True)
        df1 = df1.append(df2)
    except:
        df2 = pd.DataFrame({df1.columns[0]:0.55, df1.columns[1]:'Sin Cambios', df1.columns[2]:'Sin Cambios', df1.columns[3]:0, df1.columns[4]:0}, index=['Básico Convenio Colectivo N° 653/12 - Categoría "Oficial de Espacios Verdes"'])
        df1 = df1.append(df2)
    
    # Agrego fila "Total"
    
    df1.loc["Total"] = df1.sum()
    df1.at['Total', df1.columns[1]] = '-'
    df1.at['Total', df1.columns[2]] = '-'
    df1.at['Total', df1.columns[3]] = '-'
    
    return df1


def send_email(send_to, subject, df):
    send_from = "riquelmesebas412@gmail.com"
    password = "wamm vzam eskg gudb"
    message = """    <p><strong>Redeterminación de precios&nbsp;</strong></p>
    <p><br></p>
    <p><strong>Saludos&nbsp;</strong><br><strong>Esteban&nbsp;    </strong></p>
    """
    for receiver in send_to:
        multipart = MIMEMultipart()
        multipart["From"] = send_from
        multipart["To"] = receiver
        multipart["Subject"] = subject  
        attachment = MIMEApplication(df.to_csv())
        attachment["Content-Disposition"] = 'attachment; filename=" {}"'.format(f"{subject}.csv")
        multipart.attach(attachment)
        multipart.attach(MIMEText(message, "html"))
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(multipart["From"], password)
        server.sendmail(multipart["From"], multipart["To"], multipart.as_string())
        server.quit()



if datetime.now().day == 22:
    print("probando, probando")
    # df = redeterminacion()
    # send_email(["masuelliesteban@gmail.com"], "Redeterminacion de precios " + df.columns[2], df)
