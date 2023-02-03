# Lib para tratamento e manipulação de dados
import pandas as pd
import numpy as np

# Lib para criação do web app
import streamlit as st

# Lib para conexão para o banco de dados
import sqlite3

# Lib para importação do modelo
import joblib

# Lib para manipulação de data
from datetime import datetime

# Lib para sistema operacional
import os
import platform
import win32com.client as win32

# Lib para conexão
import socket

# Lib para criptografia
from hashlib import sha256

# Configurando o nome da aba na web
st.set_page_config(page_title = 'Aplicação Attrition')

# Importando o modelo
model = joblib.load('attrition.pkl')

def homepage():

    # Título da página
    st.title('Modelo de machine learning sobre attrition\n')
    
    # Inserindo um subtítulo
    st.header('Objetivo:')

    # Inserindo os textos
    st.markdown('Esse modelo tem como objetivo prever a probabilidade de um colaborador pedir demissão.')
    st.markdown('Para construir este modelo, usamos uma base de attrition do kaggle que foi disponibilizada pela IBM.')
    st.markdown('A importância de ter mapeado as possíveis demissões involuntárias, é que caso haja um colaborador\
        chave para empresa, e ele possua uma alta probabilidade de pedir demissões, podemos criar incentivos\
            para que ele não saia. Além disso, também podemos reduzir custos com encargos trabalhista, \
                uma vez que, poderíamos demitir um colaborador que já pretende sair.')

def modelo():

    # Inserindo um subtítulo na página
    st.header('Resultados do Modelo')

    # Inserindo o texto
    st.markdown('O modelo utilizado foi um LGBMClassifier.')
    st.markdown('Este modelo é baseado em arvores, assim como o decesion tree e o random forest.')
    st.markdown('Para atingir o resultado esperado, otimizamos alguns parâmetros, foram eles:')
    st.markdown('1. Num leaves: 5')
    st.markdown('2. Min child weight: 14')
    st.markdown('3. Min child sampes: 20')
    st.markdown('4. Max depth: 3')
    st.markdown('5. Scaler: RobustScaler')
    st.markdown('A métrica otimizado foi o precision, uma vez que, o ideal para nossa aplicação é acertar o maior número\
         de colaboradores que irão pedir demissão. Conseguimos atingir um precision de 0.91 nos dados de treino e 0.86 nos dados de teste.')
    st.markdown('Nossa base de treino possui as seguintes colunas : Age, BusinessTravel, DailyRate, Department\
        , DistanceFromHome, Education, EducationField, EnvironmentSatisfaction, Gender, HourlyRate, \
            JobInvolvement, JobLevel, JobRole, JobSatisfaction, MaritalStatus, MonthlyIncome, MonthlyRate,\
                NumCompaniesWorked, OverTime, PercentSalaryHike, PerformanceRating, RelationshipSatisfaction,\
                     StockOptionLevel, TotalWorkingYears, TrainingTimesLastYear, WorkLifeBalance, YearsAtCompany,\
                         YearsInCurrentRole, YearsSinceLastPromotion, YearsWithCurrManager.')

def consulta_individual(email):
    
    # Inserindo um subtítulo na página
    st.header('Nesta página você pode prever a probabilidade de um único colaborador')

    # Solicitando a idade
    idade = st.number_input('Digite sua idade:', min_value  = 15, max_value = 130)
    st.text(f'Você digitou: {idade}.')

    # Solicitando a frenquência que viagens a trabalho
    frequencia = ['Raramente', 'Frequentemente', 'Não Viaja']
    frequencia = st.selectbox('Digite o seu BusinessTravel: ', frequencia)
    st.text(f'Você digitou: {frequencia}.')

    if frequencia == 'Raramente':
        frequencia = 'Travel_Rarely'

    elif frequencia == 'Frequentemente':
        frequencia = 'Travel_Frequently'

    else:
        frequencia = 'Non-Travel'

    # Solicitando o DailyRate
    daily_rate = st.number_input('Digite seu DailyRate:', min_value  = 100, max_value = 1500)
    st.text(f'Você digitou: {daily_rate}.')

    # Solicitando o Department
    departamento = ['Sales', 'Research & Development', 'Human Resources']
    departamento = st.selectbox('Digite o seu Department: ', departamento)
    st.text(f'Você digitou: {departamento}.')
    
    # Solicitando a DistanceFromHome
    distance = st.number_input('Digite seu DistanceFromHome:', min_value  = 0, max_value = 30)
    st.text(f'Você digitou: {distance}.')
    
    # Solicitando a Education
    education = st.number_input('Digite seu Education:', min_value  = 1, max_value = 5)
    st.text(f'Você digitou: {education}.')

    # Solicitando a EducationField
    educationfield = ['Life Sciences', 'Other', 'Medical', 'Marketing', 'Technical Degree', 'Human Resources']
    educationfield = st.selectbox('Digite o seu EducationField: ', educationfield)
    st.text(f'Você digitou: {educationfield}.')

    # Solicitando a EnvironmentSatisfaction
    environmentsatisfaction = st.number_input('Digite seu EnvironmentSatisfaction:', min_value  = 1, max_value = 4)
    st.text(f'Você digitou: {environmentsatisfaction}.')
    
    # Solicitando a Gender
    gender = ['Female', 'Male']
    gender = st.selectbox('Digite o seu Gender: ', gender)
    st.text(f'Você digitou: {gender}.')

    if gender == "Male":
        gender = 1
    
    else:
        gender = 0

    # Solicitando a HourlyRate
    hourlyrate = st.number_input('Digite seu HourlyRate:', min_value  = 30, max_value = 100)
    st.text(f'Você digitou: {hourlyrate}.')
    
    # Solicitando a JobInvolvement
    jobninvolvement = st.number_input('Digite seu JobInvolvement:', min_value  = 1, max_value = 4)
    st.text(f'Você digitou: {jobninvolvement}.')
    
    # Solicitando a JobLevel
    joblevel = st.number_input('Digite seu JobLevel:', min_value  = 1, max_value = 5)
    st.text(f'Você digitou: {joblevel}.')
    
    # Solicitando a JobRole
    jobrole = ['Sales Executive', 'Research Scientist', 'Laboratory Technician',
       'Manufacturing Director', 'Healthcare Representative', 'Manager',
       'Sales Representative', 'Research Director', 'Human Resources']
    jobrole = st.selectbox('Digite o seu JobRole: ', jobrole)
    st.text(f'Você digitou: {jobrole}.')
    
    # Solicitando a JobSatisfaction
    jobsatisfaction = st.number_input('Digite seu JobSatisfaction:', min_value  = 1, max_value = 4)
    st.text(f'Você digitou: {jobsatisfaction}.')
    
    # Solicitando a MaritalStatus
    maritalstatus = ['Single', 'Married', 'Divorced']
    maritalstatus = st.selectbox('Digite o seu MaritalStatus: ', maritalstatus)
    st.text(f'Você digitou: {maritalstatus}.')
    
    # Solicitando a MonthlyIncome
    monthlyincome = st.number_input('Digite seu MonthlyIncome:', min_value  = 1009, max_value = 30000)
    st.text(f'Você digitou: {monthlyincome}.')
    
    # Solicitando a MonthlyRate
    monthlyrate = st.number_input('Digite seu MonthlyRate:', min_value  = 2094, max_value = 30000)
    st.text(f'Você digitou: {monthlyrate}.')
    
    # Solicitando a NumCompaniesWorked
    numcompaniesworked = st.number_input('Digite seu NumCompaniesWorked:', min_value  = 1, max_value = 9)
    st.text(f'Você digitou: {numcompaniesworked}.')
    
    # Solicitando a OverTime
    overtime = ['Yes', 'No']
    overtime = st.selectbox('Digite o seu OverTime: ', overtime)
    st.text(f'Você digitou: {overtime}.')

    if overtime == "Yes":
        overtime = 1
    
    else:
        overtime = 0
    
    # Solicitando a PercentSalaryHike
    percentsalaryhike = st.number_input('Digite seu PercentSalaryHike:', min_value  = 11, max_value = 25)
    st.text(f'Você digitou: {percentsalaryhike}.')
    
    # Solicitando a PerformanceRating
    performancerating = st.number_input('Digite seu PerformanceRating:', min_value  = 1, max_value = 4)
    st.text(f'Você digitou: {performancerating}.')
    
    # Solicitando a RelationshipSatisfaction
    relationshipsatisfaction = st.number_input('Digite seu RelationshipSatisfaction:', min_value  = 1, max_value = 4)
    st.text(f'Você digitou: {relationshipsatisfaction}.')
    
    # Solicitando a StockOptionLevel
    stockoptionlevel = st.number_input('Digite seu StockOptionLevel:', min_value  = 0, max_value = 3)
    st.text(f'Você digitou: {stockoptionlevel}.')
    
    # Solicitando a TotalWorkingYears
    totalworkingyears = st.number_input('Digite seu TotalWorkingYears:', min_value  = 0, max_value = 40)
    st.text(f'Você digitou: {totalworkingyears}.')
    
    # Solicitando a TrainingTimesLastYear
    trainingtimeslastyear = st.number_input('Digite seu TrainingTimesLastYear:', min_value  = 0, max_value = 6)
    st.text(f'Você digitou: {trainingtimeslastyear}.')
    
    # Solicitando a WorkLifeBalance
    worklifebalance = st.number_input('Digite seu WorkLifeBalance:', min_value  = 1, max_value = 4)
    st.text(f'Você digitou: {worklifebalance}.')
    
    # Solicitando a YearsAtCompany
    yearsatcompany = st.number_input('Digite seu YearsAtCompany:', min_value  = 0, max_value = 40)
    st.text(f'Você digitou: {yearsatcompany}.')
    
    # Solicitando a YearsInCurrentRole
    yearsincurrentrole = st.number_input('Digite seu YearsInCurrentRole:', min_value  = 0, max_value = 18)
    st.text(f'Você digitou: {yearsincurrentrole}.')
    
    # Solicitando a YearsSinceLastPromotion
    yearssincelastpromotion = st.number_input('Digite seu YearsSinceLastPromotion:', min_value  = 0, max_value = 15)
    st.text(f'Você digitou: {yearssincelastpromotion}.')
    
    # Solicitando a YearsWithCurrManager
    yearswithcurrmanager = st.number_input('Digite seu YearsWithCurrManager:', min_value  = 0, max_value = 17)
    st.text(f'Você digitou: {yearswithcurrmanager}.')

    # Criando uma lista com a resposta do usuário
    lista = [
        idade,
        frequencia,
        daily_rate,
        departamento,
        distance,
        education,
        educationfield,
        environmentsatisfaction,
        gender,
        hourlyrate,
        jobninvolvement,
        joblevel,
        jobrole,
        jobsatisfaction,
        maritalstatus,
        monthlyincome,
        monthlyrate,
        numcompaniesworked,
        overtime,
        percentsalaryhike,
        performancerating,
        relationshipsatisfaction,
        stockoptionlevel,
        totalworkingyears,
        trainingtimeslastyear,
        worklifebalance,
        yearsatcompany,
        yearsincurrentrole,
        yearssincelastpromotion,
        yearswithcurrmanager
    ]

    # Criando um dataframe com as colunas do modelo
    df = pd.DataFrame(columns = ['Age', 'BusinessTravel', 'DailyRate', 'Department', 'DistanceFromHome',
       'Education', 'EducationField', 'EnvironmentSatisfaction', 'Gender',
       'HourlyRate', 'JobInvolvement', 'JobLevel', 'JobRole',
       'JobSatisfaction', 'MaritalStatus', 'MonthlyIncome', 'MonthlyRate',
       'NumCompaniesWorked', 'OverTime', 'PercentSalaryHike',
       'PerformanceRating', 'RelationshipSatisfaction', 'StockOptionLevel',
       'TotalWorkingYears', 'TrainingTimesLastYear', 'WorkLifeBalance',
       'YearsAtCompany', 'YearsInCurrentRole', 'YearsSinceLastPromotion',
       'YearsWithCurrManager'])

    # Inserindo a resposta do usuário no dataframe
    df.loc[0] = lista

    # Data e hora de inicio
    data_inicio = datetime.now()

    if st.button('Calcular previsão: '):

        # Data e hora de inicio da previsão
        data_inicio = datetime.now()
        try:
            # Prevendo os valores
            prev = model.predict(df)

            # Prevendo a probabilidade
            prev_proba = model.predict_proba(df)
            st.success(f'A previsão do attrition foi de {prev[0]} com uma probabilidade de {round(prev_proba[0][1] * 100, 2)}% de pedir demissão.')
        except:
            st.warning('Não foi possível realizar a previsão, confirar as informações inserida anteriormente.')

        # Termino do processo
        data_fim = datetime.now()
        processamento = data_fim - data_inicio

        # Identificando o usuário
        usuario = os.environ.get('USERNAME')

        # Identificando o computador
        computador = platform.node()

        # Identificando o sistema operacional
        sistema = platform.platform()

        # Identificando o IP
        ip = socket.gethostbyname( socket.gethostname() )

        # Adicionando as informações no dataframe
        df['Attrition'] = prev[0]
        df['Data Inicio']  = data_inicio
        df['Data Fim']  = data_fim
        df['Delta Tempo']  = processamento
        df['E-mail'] = email
        df['Usuario'] = usuario
        df['Maquina'] = computador
        df['Sistema Operacional'] = sistema
        df['IP'] = ip

        # --------- CONEXÃO BANCO DE DADOS ---------
        #Criar a conexão com o banco de dados
        conexao_banco = sqlite3.connect('BD_Attrition.db')

        # Registrando a consulta no banco de dados
        df.to_sql(
            name = 'consulta',
            con = conexao_banco,
            if_exists = 'append',
            index  = False
        )
        conexao_banco.close()

def consulta_base(email):

    # Inserindo um subtítulo na página
    st.header('Nesta página você pode consultar diversos colaboradores')

    # Solicitando um arquivo
    xlsx = st.file_uploader("Insira um arquivo xlsx")
    if xlsx is not None:
        xlsx = pd.read_excel(xlsx)
    
    if st.button('Calcular Previsão'):
        # Data e hora de inicio
        data_inicio = datetime.now()
        try:
            # Prevendo a base
            prev_xlsx = model.predict(xlsx)

            # Anexando o resultado
            xlsx['Attrition'] =  prev_xlsx
            st.success(f'Sua previsão foi realizada!')
        except:
            st.warning('Algo deu errado, revise sua base de dados.')

        # Termino do processo
        data_fim = datetime.now()
        processamento = data_fim - data_inicio

        # Identificando o usuário
        usuario = os.environ.get('USERNAME')

        # Identificando o computador
        computador = platform.node()

        # Identificando o sistema operacional
        sistema = platform.platform()

        # Identificando o IP
        ip = socket.gethostbyname( socket.gethostname() )

        # Copiando o arquivo
        xlsx2 = xlsx.copy()

        # Adicionando as informações no dataframe
        xlsx2['Data Inicio']  = data_inicio
        xlsx2['Data Fim']  = data_fim
        xlsx2['Delta Tempo']  = processamento
        xlsx2['E-mail'] = email
        xlsx2['Usuario'] = usuario
        xlsx2['Maquina'] = computador
        xlsx2['Sistema Operacional'] = sistema
        xlsx2['IP'] = ip

        # --------- CONEXÃO BANCO DE DADOS ---------
        #Criar a conexão com o banco de dados
        conexao_banco = sqlite3.connect('BD_Attrition.db')

        # Registrando a consulta no banco de dados
        xlsx2.to_sql(
            name = 'consulta',
            con = conexao_banco,
            if_exists = 'append',
            index  = False
        )
        conexao_banco.close()

    @st.cache
    def convert_df(data):
        # IMPORTANT: Cache the conversion to prevent computation on every rerun
        return data.to_csv().encode('utf-8')
   
    xlsx = convert_df(xlsx)
    
    # Criando um botão de download da previsão
    st.download_button(
        label = 'Baixe sua previsão',
        data = xlsx,
        file_name = 'previsao.csv',
        mime = 'text/csv'
    )

        
def main():
    options = ['Login', 'Cria sua conta']

    page_option = st.sidebar.selectbox('Páginas', options)

    if page_option == 'Login':
        # Solicitando o login
        st.header('Faça o login.')
            
        # Solicitando o e-mail de login
        username = st.sidebar.text_input('Digite seu e-mail:')

        #Solicitando a senha de login
        senha = st.sidebar.text_input('Digite sua senha:', type='password')
        senha_cod = sha256(senha.encode()).hexdigest()
        
        # --------- CONEXÃO BANCO DE DADOS ---------
        #Criar a conexão com o banco de dados
        try:
            conexao_banco = sqlite3.connect('BD_Attrition.db')
            
            # Consultando a tabela
            df_login = pd.read_sql_query("""
            SELECT * FROM dados_login
            """, conexao_banco)
        except:
            pass

        # Conferindo se o e-mail e a senha estão na nossa base
        if st.sidebar.checkbox('Login'):
            df_verifica = df_login.loc[ df_login['E-mail'] == username]
            if senha_cod in df_verifica['Senha'].to_list():
                st.success('Login realizado com sucesso!')
                conexao_banco.close()
                options2 = ['Página Inicial', 'Modelo', 'Individual', 'Base de Dados']
                page_option2 = st.selectbox('Páginas', options2)
                if page_option2 == 'Página Inicial':
                    homepage()
                elif page_option2 == 'Modelo':
                    modelo()
                elif page_option2 == 'Individual':
                    consulta_individual(username)
                elif page_option2 == 'Base de Dados':
                    consulta_base(username)
                
            else:
                st.error('Dados de login inválido.')
                conexao_banco.close()

    else:
        novo_email = st.text_input('Digite o e-mail para sua conta:')
        nova_senha = st.text_input('Digite uma senha para sua conta:', type='password')
        nova_senha_cod = sha256(nova_senha.encode()).hexdigest()
        df_login = pd.DataFrame(columns=['E-mail', 'Senha'])
        df_login.loc[0] = [novo_email, nova_senha_cod]
        if st.button('Criar'):
            # --------- CONEXÃO BANCO DE DADOS ---------
            #Criar a conexão com o banco de dados
            conexao_banco = sqlite3.connect('BD_Attrition.db')
            df_login.to_sql(
                name = 'dados_login',
                con = conexao_banco,
                if_exists = 'append',
                index  = False
            )
            conexao_banco.close()
            st.success('Conta criado com sucesso!')
            st.info('Volte para página de login.')

if __name__ == '__main__':
    main()