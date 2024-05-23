#Pacotes nescessários instalar: Pandas, xlsxwriter, flask, pdfkit, requests, docxtpl, wkhtmltopdf(programa)
from flask import Flask, redirect, url_for, render_template, request, session, send_file
from datetime import timedelta
from BancoDados import Empresa
import pandas as pd
import time
from time import sleep
import math
import numpy as np
import os
import requests
from docxtpl import DocxTemplate as Dc
import pdfkit
import subprocess
from Cliente import Cliente
import cerebro
import random
import operator
import logging
import numpy as np
from functools import reduce

os.environ["TZ"] = "America/Sao_Paulo"
time.tzset()

def Tempo(Dia,Mes,Hora,Minuto):
    ano = int(time.strftime("%Y",time.localtime()))
    mes = int(time.strftime("%m",time.localtime()))
    time_string = f"{Dia} {Mes} {ano} {Hora} {Minuto}"
    seconds = time.strptime(time_string, "%d %m %Y %H %M")
    return seconds,time.mktime(seconds)

def TempoTextoAgenda(Segundos,segundosFuturo):
    dias_semana = {"Mon":"Segunda","Tue":"Terça","Wed":"Quarta","Thu":"Quinta","Fri":"Sexta","Sat":"Sábado","Sun":"Domingo"}
    tempo1 = time.localtime(Segundos)
    tempo2 = time.localtime(Segundos+segundosFuturo)
    #["Segunda","João","Pedro","13:00-14:00",dicionario_cores["Dark"]]
    return dias_semana[time.strftime('%a', tempo1)]+" "+time.strftime('%d/%m', tempo1), time.strftime("%H:%M", tempo1)+"-"+time.strftime("%H:%M", tempo2)


cliente = Cliente()
empresa = Empresa()

# 1. Obtém o endereço IP do servidor
#ipconfig_output = subprocess.check_output("ipconfig", encoding='ISO-8859-1').replace('\x87','ç').replace(' ','').replace('...............','').split("\n")
#print(ipconfig_output)
#ip_address = '192.168.15.101'

# 2. Abre a porta 5000 no firewall do Windows
#subprocess.call(["netsh", "advfirewall", "firewall", "add", "rule", "name=Flask", "dir=in", "action=allow", "protocol=TCP", "localport=5000"])

# 3. Obtém o endereço IP público
#public_ip_address = subprocess.check_output(["nslookup", "myip.opendns.com.", "resolver1.opendns.com"]).decode().split("\n")[1].strip()

# Imprime o endereço IP público
#print("Endereço IP público: ", public_ip_address)


#Variáveis Globais
binario = [["Não"],["Sim"]]

#Variáveis locais pagina coleteLeads
ninchos_atuantes = [["Educação"],["Saúde"],["Academia"],["Salão de beleza"],["Barbearia"],["Personal"],["Roupas"],["Roupas e calçados"],["Dança"]]
sub_ninchos=[["Adulto"],["Infantil"],["Sênior"],["Feminino"],["Masculino"],["Família"]]
colunas_coleta_leads = ["Nome da empresa","Nome do cliente","Link google maps","Contato1","Contato2","Data de entrada","Data pri contato","Status da primeira ligação","Data seg contato","Status da segunda ligação","Data ter contato","Status da terceira ligação","Data da visita","horario","Endereço","Nicho atuante","Sub-nicho","Tem site","Bairro","Preço m2","m2","Quantidade de equipamentos","Quantidade de funcionários","Fluxo de clientes","Anos no mercado","Cidade","Especulação de faturamento","Rebranding","Obs entrevista"]
area_empresa = [["40 m²"],["50 m²"],["60 m²"],["70 m²"],["80 m²"],["90 m²"],["100 m²"],["110 m²"],["120 m²"],["130 m²"],["140 m²"],["150 m²"],["160 m²"],["170 m²"],["180 m²"],["190 m²"],["200 m²"]]
quantidade_funcionarios = [["1"],["2"],["3"],["4"],["5"],["6"]]
fluxo_clientes = {"Pouco Movimentado":1,"Normal":2,"Movimentado":3}
anos_mercado = [["1"],["2"],["3"],["4"],["5"],["6"],["7"],["8"],["9"],["10"]]
cidades_atuantes = [["São Paulo"],["Curitiba"]]
valor_metro_bairros = ["5","10","15","20","25","30","35","40","45","50","55","60"]

#Variáveis de tempo
minutos=["00","05","10","15","20","25","30","35","40","45","50","55","60"]
horas= ["0"," 1"," 2"," 3"," 4"," 5"," 6"," 7"," 8"," 9","10","11","12","13","14","15","16","17","18","19","20","21","22","23"]
dias = ["01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"]
meses = ["01","02","03","04","05","06","07","08","09","10","11","12"]
anos = ["2023","2024"]

dicionario_cores = {"Azul":"primary","Cinza":"secondary","Verde":"success","Vermelho":"danger","Amarelo":"warning","Azul Claro":"info","Dark":"dark",}
dicionario_meses = {"Janeiro":[1,31], "Fevereiro":[2,29], "Março":[3,31], "Abril":[4,30], "Maio":[5,31], "Junho":[6,30], "Julho":[7,31], "Agosto":[8,31], "Setembro":[9,30], "Outubro":[10,31], "Novembro":[11,30], "Dezembro":[12,31]}

#tabela de valores
valor_pacotes_site=[3159,1206,3084,997.90]

#iniciando app flask na porta 5000
app = Flask(__name__)
porta = 5000
#texto de criptografia do metodo POST
app.secret_key = "textoSecreto"
#tempo de permanencia da sessão
app.permanent_session_lifetime = timedelta(days=5)





#Rotas de Login e Logout
#rota de acesso /login
#funcões: Autenticar usuário
@app.route("/login", methods=["POST","GET"])
def login():
    #verifica se usuário está logado
    if "empresa" in session:
        #carrega usuario logado
        #redireciona para a tela home
        if session["nivel"]<2:
            return redirect(url_for("home"))
        else:
            return redirect(url_for("homeusuario"))
    else:
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            #carrega dados do usuário preenchidos na tela de login
            usuario = request.form["usuario"]
            senha = request.form["senha"]
            #carrega banco de dados dos usuários
            user = Empresa().RetornarItemString("tbl_usuarios","login",usuario)
            if user == None:
                return render_template("usuario.html",erro="Usuario não encontrado")

            #verifica se usuário fornecido está no banco de dados
            if user[2]==senha:
                #torna a sessão "permanente" (configurado anteriormente para 5 dias)
                #print
                session.permanent = 1
                #carrega usuario na sessão
                session["usuario"] = user[1]
                session["empresa"] = user[3]
                session["nivel"] = user[4]
                #redireciona para a tela home
                if user[4]<2:
                    return redirect(url_for("home"))
                elif user[3] =="FOCO EDUCAÇÃO":
                    return redirect(url_for("homeusuario"))
            else:
                #retorna mensagem de erro nas credênciais
                msg = "Senha incorreta"
                #carrega tela de login
                return render_template("usuario.html",erro=msg)
        else:
            #carrega tela de login
            return render_template("usuario.html")

@app.route("/logout")
def logout():
    session.pop("usuario", None)
    session.pop("empresa", None)
    session.pop("nivel", None)
    return redirect(url_for("login"))









"""
HomePage
Data de criação: 28/12/2022
Objetivo: Mostrar os últimos agendamentos por em ordem cronológica com variação de cores para as diferentes urgências do atendimento.
Também houve uma melhora do nav-bar, para menus dropdown.
"""

#rota de acesso /
#informações mostradas: clientes à mais tempo sem serem atualizados e os prazos próximos
@app.route("/", methods=["POST","GET"])
def home():
    #verifica se usuário está logado
    if ("empresa" in session):
        #carrega usuario logado
        user = session["empresa"]
        nivel = session["nivel"]
        if nivel<2:
            pass
        else:
            return redirect(url_for("homeusuario"))
        #carrega banco de dados do usuario em questão
        clientes = Empresa().CarregarLeads(user)
        #lista aqueles a mais tempo não modificados
        #lista daqueles com o prazo terminando em um dia
        #prazo_vermelho = clientes[(clientes['Data da visita'] >= formatoData(0)[0:18]) & (clientes['Data da visita'] < formatoData(2)[0:18]) & (clientes['Data da visita'].notnull())]
        prazo_vermelho = clientes
        #lista daqueles com prazo entre 1 e 5 dias
        #prazo_amarelo = clientes[(clientes['Data da visita'] > formatoData(2)[0:18]) & (clientes['Data da visita'] < formatoData(6)[0:18]) & (clientes['Data da visita'].notnull())]
        prazo_amarelo = clientes

        #carrega html da rota /
        return render_template("homepage.html", tabela_modificado=clientes,tabela_prazo_vermelho= prazo_vermelho,tabela_prazo_amarelo= prazo_amarelo)
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))





#Função: Adicionar novos clientes
@app.route("/sistema/novocliente", methods=["POST","GET"])
def novocliente():
    if ("empresa" in session) and (session["nivel"]<2):
        user = session["empresa"]
        if request.method == "POST":
            user = Empresa().RetornarItemString("tbl_usuarios","login",request.form["login"])
            if user == None:
                Empresa().NovoDado("tbl_usuarios","login,senha,empresa,nivel",f'"{request.form["login"]}","{request.form["senha"]}","{request.form["empresa"]}",{request.form["nivel"]}')
                return render_template("novocliente.html", erro = "Cliente adicionado")
            else:
                return render_template("novocliente.html", erro = "Erro! Usuario já existe!")
        else:
            return render_template("novocliente.html")
    else:
        return redirect(url_for("login"))








'''
Sessão de paginas do cliente Foco educação
'''
#rota de acesso /usuario
#informações mostradas: clientes à mais tempo sem serem atualizados e os prazos próximos
@app.route("/usuario", methods=["POST","GET"])
def homeusuario():
    #verifica se usuário está logado
    if "empresa" in session:
        if request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Confirmar':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'CONFIRMADO'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))  
        elif request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Cancelar':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'CANCELADO'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))  
        elif request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Comparecer':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'COMPARECEU'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))           
        agenda = []
        dias_semana = []
        for i in range(7):
            dias_semana_item , horario = TempoTextoAgenda(time.time()+24*60*60*i,60)
            dias_semana.append(dias_semana_item)
        estrutura,segundos = Tempo(int(time.strftime("%d",time.localtime())),int(time.strftime("%m",time.localtime())),1,0)
        aulas = Empresa().RetornarDadosCondicao("foco_aulas","*", f"entrada > {segundos} AND entrada < {segundos+7*24*60*60}" )
        for i in aulas:
            if session["nivel"]==4:
                if i[1] == session["usuario"]:
                    dia_semana_item, horario=TempoTextoAgenda(i[3],i[4])
                    agenda.append([dia_semana_item,i[1],i[2],horario,Empresa().RetornarDadosCondicao("foco_professores","cor",f"nome = '{i[1]}'")[0][0],i[6],i[5],i[0]])
            else:
                dia_semana_item, horario=TempoTextoAgenda(i[3],i[4])
                agenda.append([dia_semana_item,i[1],i[2],horario,Empresa().RetornarDadosCondicao("foco_professores","cor",f"nome = '{i[1]}'")[0][0],i[6],i[5],i[0]])

        if session["nivel"]==4:
            professores = session["usuario"]
        else:
            professores = Empresa().RetornarDados("foco_professores","nome")
        #ans = reduce(lambda re, x: re+[x] if x not in re else re, np.array(sorted_list).transpose()[1], [])
        #logging.warning(ans)
        #carrega html da rota /usuario
        return render_template("homepageusuario.html",datas = dias_semana, tabela = agenda, professores=professores,nivel=session["nivel"])
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/agenda/add_agenda", methods=["POST","GET"])
def AgendaAdd():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<4):
        duracao = {"1h":60*60,"1h30m":90*60,"2h":2*60*60}
        materias = ["Física","Língua Portuguesa","Matemática","Biologia","Química","Língua Inglesa","Filosofia e Sociologia","Geografia","História","Ciências","Alfabetização","Informática","Terapia","Reunião","Psicopedagogia"]
        cores = [i for i in dicionario_cores.keys()]
        professores = Empresa().RetornarDadosComplemento("foco_professores","nome",condicao="ORDER BY nome")
        alunos = Empresa().RetornarDadosComplemento("foco_alunos","nome",condicao="ORDER BY nome")

        if request.method == "POST":
            estrutura, segundos = Tempo(int(request.form["dia"]),int(request.form["mes"]),int(request.form["hora"]),int(request.form["min"]))
            professor = request.form["professores"]
            aulas = Empresa().RetornarDadosCondicao("foco_aulas","*", f"entrada > {segundos-3*60*60} AND entrada < {segundos+5*60*60} AND professor = '{professor}'" )
            duracao_aula = duracao[request.form["duracao"]]
            verificador = True
            for i in aulas:
                if i[6]!="CANCELADO":
                    if segundos <= i[3]:
                        if segundos+duracao_aula > i[3]:
                            verificador = False
                            break
                    elif segundos > i[3]:
                        if segundos < i[3]+i[4]:
                            verificador = False
                            break

            if verificador:
                Empresa().NovoDado("foco_aulas","professor,aluno,entrada,saida,materia,confirmacao",f'"{professor}","{request.form["alunos"]}",{segundos},{duracao_aula},"{request.form["materias"]}","NÃO CONFIRMADO"')
                return redirect(url_for("homeusuario"))
            else:
                return render_template("Foco_agenda_add.html",erro=f"Conflito de Horário",professores = professores,alunos = alunos,materias = materias,duracao = [i for i in duracao.keys()], dia=[i for i in range(1,32)], mes=[i for i in range(1,13)], ano=[int(time.strftime("%Y",time.localtime()))+i for i in range(1)], hora=[i for i in range(25)], min=[30*i for i in range(2)])
        else:
            return render_template("Foco_agenda_add.html",professores = professores,alunos = alunos,materias = materias,duracao = [i for i in duracao.keys()], dia=[i for i in range(1,32)], mes=[i for i in range(1,13)], ano=[int(time.strftime("%Y",time.localtime()))+i for i in range(1)], hora=[i for i in range(25)], min=[30*i for i in range(2)])
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/agenda/escolher_mes_agenda", methods=["POST","GET"])
def AgendaEscolherMes():
    #verifica se usuário está logado
    if "empresa" in session:
        if request.method == "POST":
            session["mes_agenda"] = dicionario_meses[request.form["mes"]][0]
            session["dias_mes_agenda"] = dicionario_meses[request.form["mes"]][1]
            return redirect(url_for("AgendaMes"))
        else:
            return render_template("Foco_agenda_escolher_mes.html", mes=dicionario_meses.keys(), ano=[int(time.strftime("%Y",time.localtime()))+i for i in range(1)])
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/agenda/escolher_mes_agenda/foco_agenda_mes", methods=["POST","GET"])
def AgendaMes():
    #verifica se usuário está logado
    if "empresa" in session:
        if request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Confirmar':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'CONFIRMADO'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))
        elif request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Cancelar':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'CANCELADO'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))  
        elif request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Comparecer':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'COMPARECEU'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))
        mes_escolhido = session["mes_agenda"]
        estrutura, segundos = Tempo(1,int(mes_escolhido),2,0)
        agenda = []
        dias_semana = []
        numero_dias_mes = int(session["dias_mes_agenda"])
        for i in range(numero_dias_mes):
            dias_semana_item , horario = TempoTextoAgenda(segundos+24*60*60*i,1)
            dias_semana.append(dias_semana_item)
        #aulas = Empresa().RetornarDadosCondicao("foco_aulas","*", f"entrada > {segundos} AND entrada < {segundos+numero_dias_mes*24*60*60}" )
        aulas = Empresa().RetornarDados("foco_aulas","*")
        for i in aulas:
            if session["nivel"]==4:
                if i[1] == session["usuario"]:
                    dia_semana_item, horario=TempoTextoAgenda(i[3],i[4])
                    agenda.append([dia_semana_item,i[1],i[2],horario,Empresa().RetornarDadosCondicao("foco_professores","cor",f"nome = '{i[1]}'")[0][0],i[6],i[5],i[0]])
            else:
                dia_semana_item, horario=TempoTextoAgenda(i[3],i[4])
                agenda.append([dia_semana_item,i[1],i[2],horario,Empresa().RetornarDadosCondicao("foco_professores","cor",f"nome = '{i[1]}'")[0][0],i[6],i[5],i[0]])

        if session["nivel"]==4:
            professores = session["usuario"]
        else:
            professores = Empresa().RetornarDados("foco_professores","nome")
        #ans = reduce(lambda re, x: re+[x] if x not in re else re, np.array(sorted_list).transpose()[1], [])
        #logging.warning(ans)
        #carrega html da rota /usuario
        return render_template("homepageusuario.html",datas = dias_semana, tabela = agenda, professores=professores,nivel=session["nivel"])
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#Função: Adicionar professor na tabela
@app.route("/empresa/add_professor", methods=["POST","GET"])
def addprofessor():
    if ("empresa" in session) and (session["nivel"]<3):
        if request.method == "POST":
            Empresa().NovoDado("foco_professores","nome,cor",f'"{request.form["nomeprofessor"]}","{dicionario_cores[request.form["cor"]]}"')
            Empresa().NovoDado("tbl_usuarios","login, senha, empresa, nivel",f'"{request.form["nomeprofessor"]}", "acessofoco", "FOCO EDUCAÇÃO", 4')
            return redirect(url_for("homeusuario"))
        else:
            return render_template("Foco_addprofessor.html", cor=dicionario_cores.keys())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#Função: Adicionar professor na tabela
@app.route("/empresa/add_assistente", methods=["POST","GET"])
def addassistente():
    if ("empresa" in session) and (session["nivel"]<3):
        if request.method == "POST":
            Empresa().NovoDado("tbl_usuarios","login, senha, empresa, nivel",f'"{request.form["nomeassistente"]}", "acessofoco", "FOCO EDUCAÇÃO", 3')
            return redirect(url_for("homeusuario"))
        else:
            return render_template("Foco_addassistente.html")
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#Função: Adicionar professor na tabela
@app.route("/empresa/add_aluno", methods=["POST","GET"])
def addaluno():
    if ("empresa" in session) and (session["nivel"]<4):
        if request.method == "POST":
            Empresa().NovoDado("foco_alunos","nome,responsavel,email,telefone,cpf",f'"{request.form["nomealuno"]}","{request.form["nomeresponsavel"]}","{request.form["email"]}","{request.form["telefone"]}","{request.form["cpf"]}"')
            return redirect(url_for("homeusuario"))
        else:
            return render_template("Foco_addaluno.html")
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#Função: Adicionar avaliações de alunos
@app.route("/dados/avaliacao", methods=["POST","GET"])
def avaliacaoAluno():
    alunos = Empresa().RetornarDadosComplemento("foco_alunos","nome",condicao="ORDER BY nome")
    if request.method == "POST":
        Empresa().NovoDado("tbl_avaliacao_alunos","nome,avaliacao,data",f'"{request.form["nome"]}","{request.form["avaliacao"]}","{formatoData(-(1/8))}"')
        return render_template("pagina_mensagem.html", texto1 = "Avaliação adicionada",texto2="Obrigado.")
    else:
        return render_template("avaliacao_aluno.html",alunos = alunos)

#Função: visualizar avaliações de alunos
@app.route("/dados/listaavaliacao", methods=["POST","GET"])
def listaavaliacao():
    if ("empresa" in session) and (session["nivel"]<3):
        dados = Empresa().RetornarTabelas("avaliacao_alunos","*")
        dados.reverse()
        if dados == None:
            return render_template("lista_avaliacoes.html")
        else:
            return render_template("lista_avaliacoes.html",tabela = dados)
    else:
        return redirect(url_for("login"))

#Função: visualizar avaliações de alunos
@app.route("/empresa/ReceitaMensalEmpresaPorProfessor", methods=["POST","GET"])
def ReceitaMensalEmpresaPorProfessor():
    if ("empresa" in session) and (session["nivel"]<5):
        if session["nivel"] == 4:
            professores = [[session["usuario"]]]
        else:
            professores = Empresa().RetornarDados("foco_professores","nome")

        if request.method == "POST":
            estrutura1, segundos1 = Tempo(1,int(request.form["mesi"]),0,0)
            estrutura2, segundos2 = Tempo(1,int(request.form["mesf"]),0,0)
            
            aulas = Empresa().RetornarDadosCondicao("foco_aulas","*", f"entrada > {segundos1} AND entrada < {segundos2} AND professor = '{request.form['professor']}' AND confirmacao = 'COMPARECEU'" )
            datas={}
            for i in aulas:
                datas.update({i[3]:TempoTextoAgenda(i[3],0)[0]})
            materias = Empresa().RetornarDadosCondicao("foco_aulas","materia", f"entrada > {segundos1} AND entrada < {segundos2} AND professor = '{request.form['professor']}' AND confirmacao = 'COMPARECEU'" )
            materias = reduce(lambda re, x: re+[x] if x not in re else re, materias, [])
            return render_template("Foco_lista receita_mensal_professor.html",datas=datas,materias = materias,aula = aulas)
        else:
            return render_template("Foco_escolher_professor_receira.html",professor = professores, mes=[i for i in range(1,13)], ano=[int(time.strftime("%Y",time.localtime()))+i for i in range(1)])
    else:
        return redirect(url_for("login"))

def salvar_cliente():
    arqDiario = pd.read_excel('AULAS.xlsx', sheet_name=0)
    mes = vmes.get()
    nome = vnome.get()
    arqDiario = arqDiario[arqDiario["NOME"]== nome]
    texto = ''
    total = ''
    linha = "_________________________________________________________________________"
    texto += linha+"\n"
    for index,i in arqDiario.iterrows():
        valor=0
        try:
            valor = int(i["VALOR"])
        except:
            valor = 0

        texto += "Matéria:"+str(i["MATERIA"])+"\nDia: "+ str(i["DIA"])+"\nValor:"+str(valor)+" Reais"
        texto += "\n"+linha+"\n"

    total += "TOTAL: "+str(arqDiario["VALOR"].sum())

    documento = Dc(os.getcwd()+"/documentos/modelo.docx")
    alteracao = {'NOME':nome.upper(),'MES': mes.upper(),'TEXTO':texto.upper(),'TOTAL': total}
    documento.render(alteracao)
    documento.save(os.getcwd()+"/documentos/"+str(mes)+"_"+str(nome)+".docx")
    tk.messagebox.showinfo(title = "Atenção", message = "Arquivo Criado!")
    app.destroy()

@app.route("/agenda/escolher_mes_agenda_pagamento", methods=["POST","GET"])
def AgendaEscolherMesPagamento():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<4):
        if request.method == "POST":
            session["mes_agenda"] = request.form["mes"]
            return redirect(url_for("AgendaMesPagamento"))
        else:
            return render_template("Foco_agenda_escolher_mes.html", mes=[i for i in range(1,13)], ano=[int(time.strftime("%Y",time.localtime()))+i for i in range(1)])
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/Foco_agenda_mes_pagamentos", methods=["POST","GET"])
def AgendaMesPagamento():
    #verifica se usuário está logado
    if "empresa" in session:
        if request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Confirmar':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'CONFIRMADO'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))  
        elif request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Cancelar':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'CANCELADO'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))  
        elif request.method == "POST" and request.form['Atualizar Cadastro'].split('|')[0] == 'Comparecer':
            Empresa().EditarTabela('foco_aulas',"confirmacao = 'COMPARECEU'",f"id = '{request.form['Atualizar Cadastro'].split('|')[1]}'")
            return redirect(url_for("homeusuario"))
        mes_escolhido = session["mes_agenda"]
        estrutura, segundos = Tempo(1,int(mes_escolhido),0,0)
        agenda = []
        dias_semana = []
        for i in range(31):
            dias_semana_item , horario = TempoTextoAgenda(segundos+24*60*60*i,60)
            dias_semana.append(dias_semana_item)
        #professores = ["João Bentivi","Ana","Leo","Maria","Pedro","Gabriel"]
        aulas = Empresa().RetornarDadosCondicao("foco_aulas","*", f"entrada > {segundos} AND entrada < {segundos+31*24*60*60}" )
        
        for i in aulas:
            dia_semana_item, horario=TempoTextoAgenda(i[3],i[4])
            agenda.append([dia_semana_item,i[1],i[2],horario,Empresa().RetornarDadosCondicao("foco_professores","cor",f"nome = '{i[1]}'")[0][0],i[6],i[5],i[0]])
        for i in range(30):
            estrutura,segundos = Tempo(random.randint(23, 29),10,random.randint(14, 19),random.choice([0, 30]))
            dia_semana_item, horario=TempoTextoAgenda(segundos,random.choice([1, 2]))
            #agenda.append([dia_semana_item,random.choice(professores),random.choice(alunos),horario,dicionario_cores[random.choice([i for i in dicionario_cores.keys()])]])
        
        professores = Empresa().RetornarDados("foco_professores","nome")
        #ans = reduce(lambda re, x: re+[x] if x not in re else re, np.array(sorted_list).transpose()[1], [])
        #logging.warning(ans)
        #carrega html da rota /usuario
        return render_template("homepageusuario.html",datas = dias_semana, tabela = agenda, professores=professores)
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))











"""
Leads
Data de criação: 28/12/2022
Objetivo: Foram implementadas as funçoes:
"Coleta de leads": Página de peenchemento dos dados ultilizados para a prospecção de possíveis clientes.
"Contato ao cliente": Página para confirmação de entrevista (agendamento) ou eliminação deste lead.
"Cálculo do valor do pacote": Página para apresentação dos valores do pacote para as mais diversas formas de pagamento.
"""

#Função: Adiciona novo lead
@app.route("/leads/coletaleads", methods=["POST","GET"])
def coletaleads():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega o nome do usuário
        user = session["empresa"]

        try:
            arquivo = open('bairros.txt', 'r')
        except:
            arquivo = open('bairros.txt', 'x')
            arquivo.close()
            arquivo = open('bairros.txt', 'a')
            arquivo.write("")
            arquivo.close()
            arquivo = open('bairros.txt', 'r')

        menuBairros = sorted(arquivo.read().split(";"), key=str.lower)

        if request.method == "POST" and request.form['salvar'] == 'salvarBairro':

            arquivo = open('bairros.txt', 'a')
            arquivo.write(";"+request.form['nomeBairro'])
            arquivo.close()

            return redirect(url_for("coletaleads"))
        if request.method == "POST" and request.form['salvar'] == 'Adicionar':
            resultado = Empresa().SalvarLeads(request.form['bairro'],request.form['nome'],request.form['maps'],request.form['contato1'],request.form['contato2'],request.form['endereco'],request.form['site'],request.form['nicho'],request.form['subnicho'],request.form['m2'],request.form['quantidadefun'],request.form['fluxo'],request.form['anos'],request.form['cidade'],request.form['Rebranding'])
            if resultado:
                msg = "Cliente adicionado com sucesso!"
                return render_template("adicionarClientes.html", erro = msg,tabelaValorPreco=valor_metro_bairros,tabelaBairro=menuBairros,tabelaReb=binario,tabelaSite=binario,tabelaCidades=cidades_atuantes,tabelaAnosMercado=anos_mercado,tabelaFluxo=[[i] for i in fluxo_clientes.keys()],tabelaFuncionarios=quantidade_funcionarios,tabelaArea=area_empresa,tabela1=ninchos_atuantes,tabela2=sub_ninchos)
            else:
                msg = "Dados invalidos"
                return render_template("adicionarClientes.html", erro = msg,tabelaValorPreco=valor_metro_bairros,tabelaBairro=menuBairros,tabelaReb=binario,tabelaSite=binario,tabelaCidades=cidades_atuantes,tabelaAnosMercado=anos_mercado,tabelaFluxo=[[i] for i in fluxo_clientes.keys()],tabelaFuncionarios=quantidade_funcionarios,tabelaArea=area_empresa,tabela1=ninchos_atuantes,tabela2=sub_ninchos)
        else:
            return render_template("adicionarClientes.html",tabelaValorPreco=valor_metro_bairros,tabelaBairro=menuBairros,tabelaReb=binario,tabelaSite=binario,tabelaCidades=cidades_atuantes,tabelaAnosMercado=anos_mercado,tabelaFluxo=[[i] for i in fluxo_clientes.keys()],tabelaFuncionarios=quantidade_funcionarios,tabelaArea=area_empresa,tabela1=ninchos_atuantes,tabela2=sub_ninchos)
    else:
        return redirect(url_for("login"))

#Função: Confirmação de entrevista ou eliminação deste lead
@app.route("/leads/contato", methods=["POST","GET"])
def contato():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega usuário logado
        user = session["empresa"]

        tabela_pesquisa = Empresa().RetornarTabelas("leads","*")

        if request.method == "POST" and request.form['contato'].split("|")[0] == 'especulacao':
            i = Empresa().RetornarTabela("lead",int(request.form['contato'].split("|")[1]))
            roteiro=f'''Você é um profissional em análise de mercado. Preveja o faturamento mensal de uma empresa em São Paulo (mesmo que seja imprecisa a previsão) com base nas seguintes informações: nome da empresa, bairro, presença online, nicho de atuação, sub nicho de atuação, tamanho do negócio, número de funcionários, fluxo de clientes, anos no mercado, cidade e se a empresa passou por um rebranding. Além disso, inclua o preço médio dos produtos ou serviços oferecidos pela empresa. Com base nesses dados, escolha o melhor método de previsão de faturamento e descreva o processo que você usará para chegar à previsão final.Sua resposta deve seguir o modelo “O faturamento mensal deve estar por volta de [especulação de faturamento mensal]”. Pense bem antes de falar.

Dados da empresa alvo:
Nome: {i[2]}
Nicho de atuação: {i[8]}
Subnicho de atuação: {i[9]}
Tem site: {i[7]}
Bairro: {i[1].split("|")[0]}
Tamanho (m²): {i[10]}
Numero de funcioários: {i[12]}
Fluxo de clientes: {i[13]}
Anos de atuação: {i[14]}
Cidade: {i[15]}
Fez rebranding: {i[16]}'''
            Empresa().AtualisarBanco("lead",i[0],"especulacaoFaturamento",cerebro.Gerar(roteiro))
            #sleep(1)
            return redirect(url_for("contato"))

        if request.method == "POST" and request.form['contato'].split("|")[0] == 'Sem Contato':
            i = Empresa().RetornarTabela("lead",int(request.form['contato'].split("|")[1]))

            Empresa().AtualisarBanco("lead",i[0],"ultimoContato",request.form['contato'].split("|")[0])
            Empresa().AtualisarBanco("lead",i[0],"dataUltimoContato",formatoData(0))
            return redirect(url_for("contato"))
        elif request.method == "POST" and request.form['contato'].split("|")[0] == 'Já contratou um serviço':
            i = Empresa().RetornarTabela("lead",int(request.form['contato'].split("|")[1]))

            Empresa().AtualisarBanco("lead",i[0],"ultimoContato",request.form['contato'].split("|")[0])
            Empresa().AtualisarBanco("lead",i[0],"dataUltimoContato",formatoData(0))
            return redirect(url_for("contato"))
        elif request.method == "POST" and request.form['contato'].split("|")[0] == 'Não quis':
            i = Empresa().RetornarTabela("lead",int(request.form['contato'].split("|")[1]))

            Empresa().AtualisarBanco("lead",i[0],"ultimoContato",request.form['contato'].split("|")[0])
            Empresa().AtualisarBanco("lead",i[0],"dataUltimoContato",formatoData(0))
            return redirect(url_for("contato"))
        elif request.method == "POST" and request.form['contato'].split("|")[0] == 'Marcar':
            return redirect(url_for("entrevista",ID=int(request.form['contato'].split("|")[1])))
        else:
            #carrega a tela de pesquisa
            return render_template("pesquisa.html", tabela=tabela_pesquisa)
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route('/aguarde')
def aguarde():
    return render_template('aguarde.html')

def Esperar():
    return redirect(url_for("aguarde"))

#Função: Confirmação de entrevista ou eliminação deste lead
@app.route("/leads/dados", methods=["POST","GET"])
def dadosLeads():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega usuário logado
        user = session["empresa"]
        if request.method == "POST" and request.form['gerar'] == 'script':
            tabela_pesquisa = Empresa().RetornarTabelas("leads","*")
            for i in tabela_pesquisa:
                if len(i[19])<20:
                    roteiro=f'''Você é da empresa Luminus. Crie um roteiro de vendas para uma agência de marketing digital com o objetivo de ajudar empresas de médio porte a alcançar mais clientes online. As empresas alvo estão localizadas em São Paulo e incluem academias, lojas de roupas e acessórios femininos e masculinos, restaurantes e cafeterias. O roteiro deve abordar empresas com faturamento mínimo de 15 a 20 mil e mais de dois ou três anos de mercado que ainda não possuem presença digital. O objetivo é ajudar os proprietários de empresas que se sentem perdidos e não sabem por onde começar. Você oferece uma consultoria de marketing gratuita e avaliação de mídias socias.

Dados da empresa alvo:
Nome: {i[2]}
Nicho de atuação: {i[8]}
Subnicho de atuação: {i[9]}
Tem site: {i[7]}
Bairro: {i[1].split("|")[0]}
Tamanho (m²): {i[10]}
Numero de funcioários: {i[12]}
Fluxo de clientes: {i[13]}
Anos de atuação: {i[14]}
Cidade: {i[15]}
Fez rebranding: {i[16]}'''
                    print(roteiro+"\n\n")
                    #print(cerebro.Gerar(roteiro))
                    #Empresa().AtualisarBanco("lead",i[0],"roteiroVendas",cerebro.Gerar(roteiro))
            sleep(5)
            return redirect(url_for("contato"))
        else:
            #carrega a tela de pesquisa
            return render_template('aguarde.html')
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))


#Função: Confirmação de entrevista ou eliminação deste lead
@app.route("/leads/especularFaturamento", methods=["POST","GET"])
def especularFaturamento():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):

        #carrega usuário logado
        user = session["empresa"]

        tabela_pesquisa = Empresa().RetornarTabelas("leads","*")
        for i in tabela_pesquisa:
            if len(i[19])<20:
                roteiro=f'''Você é um profissional em análise de mercado. Preveja o faturamento mensal de uma empresa em São Paulo (mesmo que seja imprecisa a previsão) com base nas seguintes informações: nome da empresa, bairro, presença online, nicho de atuação, sub nicho de atuação, tamanho do negócio, número de funcionários, fluxo de clientes, anos no mercado, cidade e se a empresa passou por um rebranding. Além disso, inclua o preço médio dos produtos ou serviços oferecidos pela empresa. Com base nesses dados, escolha o melhor método de previsão de faturamento e descreva o processo que você usará para chegar à previsão final.Sua resposta deve seguir o modelo “O faturamento mensal deve estar por volta de [especulação de faturamento mensal]”. Pense bem antes de falar.

Dados da empresa alvo:
Nicho de atuação: {i[8]}
Subnicho de atuação: {i[9]}
Tem site: {i[7]}
Bairro: {i[1].split("|")[0]}
Tamanho (m²): {i[10]}
Numero de funcioários: {i[12]}
Fluxo de clientes: {i[13]}
Anos de atuação: {i[14]}
Cidade: {i[15]}
Fez rebranding: {i[16]}'''
                print(roteiro+"\n\n")
                #print(cerebro.Gerar(roteiro))
                Empresa().AtualisarBanco("lead",i[0],"especulacaoFaturamento",cerebro.Gerar(roteiro))
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            return redirect(url_for("contato"))
        else:
            #carrega a tela de pesquisa
            return redirect(url_for("contato"))
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))








"""
Caminho referente ao contato
"""

#Adiciona os dados da entrevista
@app.route("/leads/andamentos/entrevista", methods=["POST","GET"])
def entrevista():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        ID = request.args['ID']
        #carrega o nome do usuário
        user = session["empresa"]
        #carrega coluna de itens que teram que selecionar na próxima página.
        item = Empresa().RetornarTabela("lead",int(ID))

        menuMinutos = pd.DataFrame(minutos)
        menuHoras = pd.DataFrame(horas)
        menuDias = pd.DataFrame(dias)
        menuMeses = pd.DataFrame(meses)
        menuAnos = pd.DataFrame(anos)
        if request.method == "POST" and request.form['acao'].split("|")[0] == 'gerarScript':
            i = Empresa().RetornarTabela("lead",int(request.form['acao'].split("|")[1]))
            roteiro=f'''Você é da empresa Luminus. Crie um roteiro de vendas para uma agência de marketing digital com o objetivo de ajudar empresas de médio porte a alcançar mais clientes online. As empresas alvo estão localizadas em São Paulo e incluem academias, lojas de roupas e acessórios femininos e masculinos, restaurantes e cafeterias. O roteiro deve abordar empresas com faturamento mínimo de 15 a 20 mil e mais de dois ou três anos de mercado que ainda não possuem presença digital. O objetivo é ajudar os proprietários de empresas que se sentem perdidos e não sabem por onde começar. Você oferece uma consultoria de marketing gratuita e avaliação de mídias socias.

Dados da empresa alvo:
Nome: {i[2]}
Nicho de atuação: {i[8]}
Subnicho de atuação: {i[9]}
Tem site: {i[7]}
Bairro: {i[1].split("|")[0]}
Tamanho (m²): {i[10]}
Numero de funcioários: {i[12]}
Fluxo de clientes: {i[13]}
Anos de atuação: {i[14]}
Cidade: {i[15]}
Fez rebranding: {i[16]}'''
            Empresa().AtualisarBanco("lead",i[0],"roteiroVendas",cerebro.Gerar(roteiro))
            return render_template("adicionarEntrevista.html",ID=ID,script=item[19],nome="",obs="",tabela1=menuMinutos.values.tolist(),tabela2=menuHoras.values.tolist(),tabela3=menuDias.values.tolist(),tabela4=menuMeses.values.tolist(),tabela5=menuAnos.values.tolist())
        else:
            return render_template("adicionarEntrevista.html",ID=ID,script=item[19],nome="",obs="",tabela1=menuMinutos.values.tolist(),tabela2=menuHoras.values.tolist(),tabela3=menuDias.values.tolist(),tabela4=menuMeses.values.tolist(),tabela5=menuAnos.values.tolist())
    else:
        return redirect(url_for("login"))

#Edita os dados da entrevista
@app.route("/leads/agendamentos/editarentrevista/<nome>", methods=["POST","GET"])
def editarEntrevista(nome):
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega o nome do usuário
        user = session["empresa"]

        clientes = carregarCliente(user)
        index = clientes[clientes['Nome da empresa'] == nome].index[0]
        #carrega coluna de itens que teram que selecionar na próxima página.
        menuMinutos = pd.DataFrame(minutos)
        menuHoras = pd.DataFrame(horas)
        menuDias = pd.DataFrame(dias)
        menuMeses = pd.DataFrame(meses)
        menuAnos = pd.DataFrame(anos)

        if request.method == "POST":
            clientes = carregarCliente(user)


            data = ConfirmarTexto("ano")+"/"+ConfirmarData("mes")+"/"+ConfirmarData("dia")+" "+ConfirmarData("hora")+":"+ConfirmarData("min")

            lugares_adicionar=["Nome do cliente","Obs entrevista","Data da visita"]
            valores_adicionar=[ConfirmarTexto("nomeCliente"),ConfirmarTexto("obs"),data]
            clientes = adicionarValorCliente("Nome da empresa",clientes,nome,lugares_adicionar,valores_adicionar)

            #index = clientes[clientes['Nome do cliente'] == nome].index[0]
            clientes.at[index,'Data ter contato'] = formatoData(0)[0:18]
            clientes.at[index,'Status da terceira ligação'] = "Agendado"

            clientes.to_excel(os.getcwd()+'/dados/clientes/clientes_luminus'+user+'.xlsx')

            msg = "Salvo com sucesso!"
            return redirect(url_for("home"))
        else:
            return render_template("adicionarEntrevista.html",nome=str(clientes.at[index,'Nome do cliente']),obs=str(clientes.at[index,'Obs entrevista']),comando = "",tabela1=menuMinutos.values.tolist(),tabela2=menuHoras.values.tolist(),tabela3=menuDias.values.tolist(),tabela4=menuMeses.values.tolist(),tabela5=menuAnos.values.tolist())
    else:
        return redirect(url_for("login"))


"""
Feramentas das paginas
"""

#Função: Adicionar a situação do contato através dos botões da página contatos.
@app.route("/leads/addandamento/<nome>/<mensagem>", methods=["POST","GET"])
def addandamento(nome,mensagem):
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega usuário logado
        user = session["empresa"]
        clientes = carregarCliente(user)
        #localiza cliente a ser modificado
        index = clientes[clientes['Nome da empresa'] == nome].index[0]
        #modifica valor no cliente
        if clientes.at[index,"Data ter contato"]=="":
            if clientes.at[index,"Data seg contato"]=="" and mensagem not in ["Já_contratou_um_serviço","Não_quis"]:
                if clientes.at[index,'Data pri contato']=="":
                    clientes.at[index,'Data pri contato'] = formatoData(0)[0:17]
                    clientes.at[index,'Status da primeira ligação'] = mensagem.replace("_"," ")
                else:
                    clientes.at[index,'Data seg contato'] = formatoData(0)[0:17]
                    clientes.at[index,'Status da segunda ligação'] = mensagem.replace("_"," ")
            else:
                clientes.at[index,'Data ter contato'] = formatoData(0)[0:17]
                clientes.at[index,'Status da terceira ligação'] = mensagem.replace("_"," ")
        clientes.to_excel(os.getcwd()+'/dados/clientes/clientes_luminus'+user+'.xlsx')
        #carrega tela de pesquisa com a lista de resultados
        return redirect(url_for("contato"))

    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))



"""
Caminho para a Primeira entrevista
"""
#Apresenta as opções de pacotes
@app.route("/projetos/primeiraentrevista", methods=["POST","GET"])
def primeiraentrevista():
    if ("empresa" in session) and (session["nivel"]<2):
        user = session["empresa"]
        if request.method == "POST":
            return redirect(url_for("perguntasentrevista",ID=request.form['cliente']))
        else:
            tabela_pesquisa = Empresa().RetornarTabelas("clientes","nome_empresa")
            return render_template("projetosEntrevistaEscolha.html",tabelaClientes=tabela_pesquisa)
    else:
        return redirect(url_for("login"))

@app.route("/projetos/perguntasentrevista", methods=["POST","GET"])
def perguntasentrevista():
    if ("empresa" in session) and (session["nivel"]<2):
        ID = request.args['ID']
        dados = Empresa().RetornarItemString("tbl_entrevista","nome_empresa",ID)
        if dados == None:
            Empresa().NovoDado("tbl_entrevista","nome_empresa,resumo,sazonalidade,mundo_perfeito,concorrentes,produto_lucrativo,rotina_geracao_audiencia,quantidades_visitantes,cliente_perfeito,informacao_importante_produto,clientes_querem,processo_vendas_atual,produtos_maior_valor,captal_giro_marketing",f'"{ID}","","","","","","","","","","","","",""')
            dados = Empresa().RetornarItemString("tbl_entrevista","nome_empresa",ID)
        user = session["empresa"]
        if request.method == "POST" and request.form['botao'] == 'Salvar':
            Empresa().EditarTabela("tbl_entrevista",f'resumo = "{request.form["resumo"]}",sazonalidade = "{request.form["sazonalidade"]}",mundo_perfeito = "{request.form["mundo_perfeito"]}",concorrentes = "{request.form["concorrentes"]}",produto_lucrativo = "{request.form["produto_lucrativo"]}",rotina_geracao_audiencia = "{request.form["rotina_geracao_audiencia"]}",quantidades_visitantes = "{request.form["quantidades_visitantes"]}",cliente_perfeito = "{request.form["cliente_perfeito"]}",informacao_importante_produto = "{request.form["informacao_importante_produto"]}",clientes_querem = "{request.form["clientes_querem"]}",processo_vendas_atual = "{request.form["processo_vendas_atual"]}",produtos_maior_valor = "{request.form["produtos_maior_valor"]}",captal_giro_marketing = "{request.form["captal_giro_marketing"]}"',f'nome_empresa = "{ID}"')
            dados = Empresa().RetornarItemString("tbl_entrevista","nome_empresa",ID)
            return render_template("infoEntrevista.html",clienteCalendario=ID,dados=dados,erro="Atualizado!")
        if request.method == "POST" and request.form['botao'] == 'Matriz Swat':
            #TODO
            return redirect(url_for("matrizswat",ID=ID))
        else:
            return render_template("infoEntrevista.html",clienteCalendario=ID,dados=dados)
    else:
        return redirect(url_for("login"))

@app.route("/projetos/perguntasentrevista/matrizswat", methods=["POST","GET"])
def matrizswat():
    if ("empresa" in session) and (session["nivel"]<2):
        ID = request.args['ID']
        dados = Empresa().RetornarItemString("tbl_matriz_swat","nome_empresa",ID)
        if dados == None:
            Empresa().NovoDado("tbl_matriz_swat","nome_empresa,fazemos_bem,diferenciais,publico_aprova,baixo_desempenho,a_melhorar,recursos_melhorar_desempenho,recursos_melhorar_fraquesas,lacunas_servicos,meta_anual,mudancas_preocupantes,tendencias_mercado,concorrentes",f'"{ID}","","","","","","","","","","","",""')
            dados = Empresa().RetornarItemString("tbl_matriz_swat","nome_empresa",ID)
        user = session["empresa"]
        if request.method == "POST" and request.form['botao'] == 'Salvar SWAT':
            Empresa().EditarTabela("tbl_matriz_swat",f'fazemos_bem = "{request.form["fazemos_bem"]}",diferenciais = "{request.form["diferenciais"]}",publico_aprova = "{request.form["publico_aprova"]}",baixo_desempenho = "{request.form["baixo_desempenho"]}",a_melhorar = "{request.form["a_melhorar"]}",recursos_melhorar_desempenho = "{request.form["recursos_melhorar_desempenho"]}",recursos_melhorar_fraquesas = "{request.form["recursos_melhorar_fraquesas"]}",lacunas_servicos = "{request.form["lacunas_servicos"]}",meta_anual = "{request.form["meta_anual"]}",mudancas_preocupantes = "{request.form["mudancas_preocupantes"]}",tendencias_mercado = "{request.form["tendencias_mercado"]}",concorrentes = "{request.form["concorrentes"]}"',f'nome_empresa = "{ID}"')
            dados = Empresa().RetornarItemString("tbl_matriz_swat","nome_empresa",ID)
            return render_template("matriz_swat.html",clienteCalendario=ID,dados=dados,erro="Atualizado!")
        if request.method == "POST" and request.form['botao'] == 'Insights':
            return redirect(url_for("insights",ID=ID,Pergunta="",Resposta=""))
        else:
            return render_template("matriz_swat.html",clienteCalendario=ID,dados=dados)
    else:
        return redirect(url_for("login"))

@app.route("/projetos/perguntasentrevista/matrizswat/insights", methods=["POST","GET"])
def insights():
    if ("empresa" in session) and (session["nivel"]<2):
        ID = request.args['ID']
        pergunta = request.args['Pergunta']
        resposta = request.args['Resposta']
        dados = Empresa().RetornarItemStrings("tbl_insights","nome_empresa",ID)
        dados.reverse()
        #    Empresa().NovoDado("tbl_insights","nome_empresa,pergunta,resposta",f'"{ID}","",""')
        #    dados = Empresa().RetornarItemString("tbl_insights","nome_empresa",ID)
        user = session["empresa"]
        if request.method == "POST" and request.form['botao'] == 'Salvar':
            pergunta = request.form["pergunta"]
            resposta = request.form["resposta"]
            Empresa().NovoDado("tbl_insights","nome_empresa,pergunta,resposta",f'"{ID}","{pergunta}","{resposta}"')
            dados = Empresa().RetornarItemStrings("tbl_insights","nome_empresa",ID)
            dados.reverse()
            return render_template("insights.html",clienteCalendario=ID,tabela = dados,pergunta="",resposta="")
        if request.method == "POST" and request.form['botao'] == 'gerar':
            dados_empresa = Empresa().RetornarItemString("tbl_entrevista","nome_empresa",ID)
            dados_matriz = Empresa().RetornarItemString("tbl_matriz_swat","nome_empresa",ID)
            pergunta = request.form['pergunta']
            roteiro = f'''Você é um profissional renomado no marketing, sua função é criar insights para perguntas de uma empresa. Abaixo seguem algumas informações sobre a empresa:

[Informações da entrevista inicial]

1- Faça um breve resumo do seu negócio:
{dados_empresa[1]}

2- Existe alguma sazonalidade no seu mercado?
{dados_empresa[2]}

3- Em um mundo perfeito, o que você gostaria que acontecesse com nosso trabalho para que você ficasse feliz?
{dados_empresa[3]}

4- Qual seu produto mais lucrativo?
{dados_empresa[5]}

5- Fale sobre sua rotina de geração de audiência e vendas:
{dados_empresa[6]}

6- Qual a quantidade de visitantes você tem por dia, semana, mês?
{dados_empresa[7]}

7- Defina em detalhes o seu cliente perfeito:
{dados_empresa[8]}

8- Qual a coisa mais importante que o seu cliente precisa saber sobre seu produto/serviço?
{dados_empresa[9]}

9- Qual a coisa mais importante que os seus clientes querem mais que qualquer outra coisa?
{dados_empresa[10]}

10- Como acontece o processo de vendas atualmente?
{dados_empresa[11]}

11- Existe algum outro produto de maior valor que pode ser oferecido?
{dados_empresa[12]}

12- Quanto você pode disponibilizar de recursos para estratégias de marketing hoje?
{dados_empresa[13]}


[Matriz SWOT]

A seguir segue informações coletadas por um profissional do marketing por meio de uma matriz SWOT.

1- Forças: O que fazemos bem?
{dados_matriz[1]}

2- Forças: O que diferencia a nossa organização?
{dados_matriz[2]}

3- Forças: De que o nosso público-alvo gosta na nossa organização?
{dados_matriz[3]}

1- Fraquezas: Quais iniciativas estão com desempenho abaixo do esperado e por quê?
{dados_matriz[4]}

2- Fraquezas: O que poderia melhorar?
{dados_matriz[5]}

3- Fraquezas: Quais recursos poderiam melhorar o nosso desempenho?
{dados_matriz[6]}

1- Oportunidades: Quais recursos podemos usar para melhorar as nossas fraquezas?
{dados_matriz[7]}

2- Oportunidades: Existem lacunas de mercado nos serviços que prestamos?
{dados_matriz[8]}

3- Oportunidades: Quais são as nossas metas para o ano?
{dados_matriz[9]}

1- Ameaças: Quais mudanças na indústria são motivos de preocupação?
{dados_matriz[10]}

2- Ameaças: Quais são as novas tendências de mercado no horizonte?
{dados_matriz[11]}

3- Ameaças: Em que pontos nossos concorrentes têm um melhor desempenho que o nosso?
{dados_matriz[12]}


[Pergunta]

A partir destas informações, responda a pergunta abaixo de maneira clara e bem analisada.

{pergunta}. Pense bem e seja claro em sua resposta.
'''

            try:
                resposta = cerebro.Gerar(roteiro)
            except:
                resposta = "Erro na criação da resposta"

            if dados == None:
                return render_template("insights.html",clienteCalendario=ID,pergunta=pergunta,resposta=resposta)
            else:
                return render_template("insights.html",clienteCalendario=ID,tabela = dados,pergunta=pergunta,resposta=resposta)
        else:
            if dados == None:
                return render_template("insights.html",clienteCalendario=ID,pergunta=pergunta,resposta=resposta)
            else:
                return render_template("insights.html",clienteCalendario=ID,tabela = dados,pergunta=pergunta,resposta=resposta)
    else:
        #return render_template("pagina_mensagem.html", texto1 = "Serviço ainda não disponível para o cliente final",texto2="Contate o programador para mais informações.")
        return redirect(url_for("login"))
    #elif session["nivel"] == "FOCO EDUCAÇÃO":
    #    return render_template("pagina_mensagem.html", texto1 = "Serviço ainda não disponível para o cliente final",texto2="Contate o programador para mais informações.")
    #else:
    #    return redirect(url_for("login"))









"""
Caminho para a criação do calendário editorial
"""
#Apresenta as opções de pacotes
@app.route("/projetos/calendarioeditorial", methods=["POST","GET"])
def calendarioeditorial():
    if ("empresa" in session) and (session["nivel"]<2):
        user = session["empresa"]
        if request.method == "POST":
            return redirect(url_for("inbound",ID=request.form['cliente']))
        else:
            tabela_pesquisa = Empresa().RetornarTabelas("clientes","nome_empresa")
            return render_template("projetosCalendarioEscolha.html",tabelaClientes=tabela_pesquisa)
    else:
        return redirect(url_for("login"))

@app.route("/projetos/inbound", methods=["POST","GET"])
def inbound():
    if ("empresa" in session) and (session["nivel"]<2):
        ID = request.args['ID']
        dados = Empresa().RetornarItemString("tbl_inbound","nome_empresa",ID)
        if dados == None:
            Empresa().NovoDado("tbl_inbound","nome_empresa,publico_alvo,persona,diferenciais,tom_de_voz,objetivos,persona_marca,palavras_preferidas",f'"{ID}","","","","","","",""')
            dados = Empresa().RetornarItemString("tbl_inbound","nome_empresa",ID)
        user = session["empresa"]
        if request.method == "POST" and request.form['botao'] == 'Salvar':
            Empresa().EditarTabela("tbl_inbound",f'publico_alvo = "{request.form["publico"]}",persona = "{request.form["persona"]}",diferenciais = "{request.form["diferenciais"]}",tom_de_voz = "{request.form["tom"]}",objetivos = "{request.form["objetivos"]}",persona_marca = "{request.form["persona_marca"]}",palavras_preferidas = "{request.form["palavras"]}"',f'nome_empresa = "{ID}"')
            dados = Empresa().RetornarItemString("tbl_inbound","nome_empresa",ID)
            return render_template("infoInbound.html",clienteCalendario=ID,dados=dados,erro="Atualizado!")
        if request.method == "POST" and request.form['botao'] == 'Gerar calendário':
            return redirect(url_for("gerarcalendario",ID=ID))
        else:
            return render_template("infoInbound.html",clienteCalendario=ID,dados=dados)
    else:
        return redirect(url_for("login"))

@app.route("/projetos/inbound/gerarcalendario", methods=["POST","GET"])
def gerarcalendario():
    if ("empresa" in session) and (session["nivel"]<2):
        ID = request.args['ID']
        dados = Empresa().RetornarItemString("tbl_calendario_publicacao","nome_empresa",ID)
        if dados == None:
            Empresa().NovoDado("tbl_calendario_publicacao","nome_empresa,temas_aprovados,conteudo_gerado,datas_comemorativas",f'"{ID}","","",""')
            dados = Empresa().RetornarItemString("tbl_calendario_publicacao","nome_empresa",ID)
        user = session["empresa"]
        if request.method == "POST" and request.form['botao'] == 'Salvar temas':
            Empresa().EditarTabela("tbl_calendario_publicacao",f'temas_aprovados = "{request.form["temas"]}",conteudo_gerado = "{request.form["gerado"]}",datas_comemorativas = "{request.form["datas"]}"',f'nome_empresa = "{ID}"')
            dados = Empresa().RetornarItemString("tbl_calendario_publicacao","nome_empresa",ID)
            return render_template("gerarcalendario.html",clienteCalendario=ID,dados=dados,erro="Atualizado!")
        if request.method == "POST" and request.form['botao'] == 'gerar':
            dados_empresa = Empresa().RetornarItemString("tbl_inbound","nome_empresa",ID)
            dados_calendario = Empresa().RetornarItemString("tbl_calendario_publicacao","nome_empresa",ID)

            roteiro = f'''Você é um profissional renomado no marketing, sua função é criar o calendário editorial para uma empresa chamada {ID}. Abaixo seguem algumas informações sobre a empresa:

            1- Público alvo:

            {dados_empresa[1]}

            2- Personas:

            {dados_empresa[2]}

            3- Diferenciais da empresa:

            {dados_empresa[3]}

            4- Tom de voz da marca:

            {dados_empresa[4]}

            5- Objetivos da empresa:

            {dados_empresa[5]}

            6- Persona da marca

            {dados_empresa[6]}

            7- Palavras ou frases preferidas

            {dados_empresa[7]}



            '''
            temas_anteriores = "\n".join(dados_calendario[1].split("\n")[-20:])

            roteiro +=f'''[USUARIO] - Você deve criar mais temas de publicações de um calendário editorial que sejam pertinentes com com a empresa acima, ultilize o padrão -[Titulo]:[Descrição]. Seja bem detalhado na descrição. Crie 15 temas inovadores, pense bem antes de escrever:

            [TEMAS DE PUBLICAÇÃO DIFERENTES E INOVADORES]
            {temas_anteriores}
            {dados_calendario[2].replace("0","").replace("1","").replace("2","").replace("3","").replace("4","").replace("5","").replace("6","").replace("7","").replace("8","").replace("9","")}

            [USUARIO] - Crie outros temas:

            '''

            '''
            else:
                roteiro +=[COMANDO] - Você deve criar mais descrições de publicações com descrições que sejam pertinentes a empresa acima, ultilize o padrão -[Titulo]:[Descrição] para cada um os titulos abaixo:

            {request.form["datas"]}

            '''
            try:
                gerado = cerebro.Gerar(roteiro)
                #gerado = "\n".join(dados_calendario[1].split("\n")[-25:])
            except:
                gerado = "Erro na criação do calendário"

            Empresa().EditarTabela("tbl_calendario_publicacao",f'conteudo_gerado = "{gerado}"',f'nome_empresa = "{ID}"')
            return redirect(url_for("gerarcalendario",ID=ID))
        else:
            return render_template("gerarcalendario.html",clienteCalendario=ID,dados=dados)
    else:
        return redirect(url_for("login"))










"""
Caminho referente ao valor
"""



#Apresenta as opções de pacotes
@app.route("/leads/valor", methods=["POST","GET"])
def valorPacote():
    if ("empresa" in session) and (session["nivel"]<2):
        user = session["empresa"]
        if request.method == "POST":
            clientes = carregarCliente(user)
            return render_template("valor_pacote.html")
        else:
            return render_template("valor_pacote.html")
    else:
        return redirect(url_for("login"))

#Apresenta os valores e opções de pagamento para dado pacote
@app.route("/leads/valor/<valor>", methods=["POST","GET"])
def CalculoValorPacote(valor):
    if ("empresa" in session) and (session["nivel"]<2):
        user = session["empresa"]
        custo_site = valor_pacotes_site[0]
        if valor == "ouro":
            custo_pacote = valor_pacotes_site[2]
            mensal = "+ Valor mensal de "+"{:.2f}".format(round(valor_pacotes_site[3]))+" R$"
        if valor == "prata":
            custo_pacote = valor_pacotes_site[2]
            mensal=""
        if valor == "bronze":
            custo_pacote = valor_pacotes_site[1]
            mensal=""

        custo_total = [custo_site,custo_pacote,custo_site+custo_pacote,round((custo_site+custo_pacote)/12,2)]
        custo_parcelado = list()
        k=0
        for i in custo_total:
            custo_parcelado.append("{:.2f}".format(round(i/0.9,2)))
            custo_total[k]= "{:.2f}".format(round(i))
            k=k+1
        preco = float(custo_parcelado[2])
        preco2 = float(custo_total[2])

        x7 = "{:.2f}".format(round(preco + preco2*0.015*1,2))+"R$ ("+"{:.2f}".format(round((preco + preco2*0.015*1)/12,2))+"R$ ao mês)"
        x8 = "{:.2f}".format(round(preco + preco2*0.015*2,2))+"R$ ("+"{:.2f}".format(round((preco + preco2*0.015*2)/12,2))+"R$ ao mês)"
        x9 = "{:.2f}".format(round(preco + preco2*0.015*3,2))+"R$ ("+"{:.2f}".format(round((preco + preco2*0.015*3)/12,2))+"R$ ao mês)"
        x10 = "{:.2f}".format(round(preco + preco2*0.015*4,2))+"R$ ("+"{:.2f}".format(round((preco + preco2*0.015*4)/12,2))+"R$ ao mês)"
        x11 = "{:.2f}".format(round(preco + preco2*0.015*5,2))+"R$ ("+"{:.2f}".format(round((preco + preco2*0.015*5)/12,2))+"R$ ao mês)"
        x12 = "{:.2f}".format(round(preco + preco2*0.015*6,2))+"R$ ("+"{:.2f}".format(round((preco + preco2*0.015*6)/12,2))+"R$ ao mês)"

        des = "{:.2f}".format(round(float(custo_parcelado[2])-float(custo_total[2])))

        if request.method == "POST":
            clientes = carregarCliente(user)
        else:
            return render_template("valor_pacote_final.html",acrec=mensal,plano = valor,desconto = des ,vista1 = custo_total[0],vista2 = custo_total[1],vista3 = custo_total[2],vista4 = custo_total[3],prazo1 = custo_parcelado[0],prazo2 = custo_parcelado[1],prazo3 = custo_parcelado[2],prazo4 = custo_parcelado[3],par7=x7,par8=x8,par9=x9,par10=x10,par11=x11,par12=x12)
    else:
        return redirect(url_for("login"))










"""
Caminho referente a reuniões
"""

#Edita os dados da entrevista
@app.route("/reuniao/nova", methods=["POST","GET"])
def novaReuniao():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega o nome do usuário
        user = session["empresa"]

        #tabelas usadas
        prioridade = pd.DataFrame(["Não Urgente","Urgente"])
        concluido = pd.DataFrame(["Não","Sim"])
        if request.method == "POST":
            dados = carregarReunioes(user)

            lista = {"Data":formatoData(0),"Participantes":ConfirmarTexto("participantes"),"Tema da reunião":ConfirmarTexto("temaReuniao"),"Resumo da reunião":ConfirmarTexto("resumo"),"Anotações":ConfirmarTexto("anotacoes"),"Lista de decisões":ConfirmarTexto("decisoes"),"Tarefas":ConfirmarTexto("tarefas"),"Prioridade":ConfirmarTexto("prioridade"),"Prazo":ConfirmarTexto("prazo"),"Concluido":ConfirmarTexto("concluido")}
            item =pd.Series(lista,index=dados.columns)
            dados.loc[len(dados)]=item
            dados.to_excel(os.getcwd()+'/dados/clientes/'+user+'/reunioes_'+user+'.xlsx')

            msg = "Salvo com sucesso!"
            return redirect(url_for("pautas"))
        else:
            return render_template("adicionarReuniao.html",tabelaPrioridade=prioridade.values.tolist(),tabelaConcluido=concluido.values.tolist())
    else:
        return redirect(url_for("login"))


#Edita os dados da entrevista
@app.route("/reuniao/editar/<data>", methods=["POST","GET"])
def editarReuniao(data):
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega o nome do usuário
        user = session["empresa"]

        dados = carregarReunioes(user)
        index = dados[dados["Data"]==data].index[0]



        #tabelas usadas
        prioridade = pd.DataFrame(["Não Urgente","Urgente"])
        concluido = pd.DataFrame(["Não","Sim"])
        if request.method == "POST":
            dados = carregarReunioes(user)

            lista = {"Data":formatoData(0),"Participantes":ConfirmarTexto("participantes"),"Tema da reunião":ConfirmarTexto("temaReuniao"),"Resumo da reunião":ConfirmarTexto("resumo"),"Anotações":ConfirmarTexto("anotacoes"),"Lista de decisões":ConfirmarTexto("decisoes"),"Tarefas":ConfirmarTexto("tarefas"),"Prioridade":ConfirmarTexto("prioridade"),"Prazo":ConfirmarTexto("prazo"),"Concluido":ConfirmarTexto("concluido")}
            item =pd.Series(lista,index=dados.columns)
            dados.loc[len(dados)]=item
            dados.to_excel(os.getcwd()+'/dados/clientes/'+user+'/reunioes_'+user+'.xlsx')

            msg = "Salvo com sucesso!"
            return redirect(url_for("pautas"))
        else:
            return render_template("adicionarReuniao.html",tabelaPrioridade=prioridade.values.tolist(),tabelaConcluido=concluido.values.tolist())
    else:
        return redirect(url_for("login"))






#rota de acesso /
#informações mostradas: clientes à mais tempo sem serem atualizados e os prazos próximos
@app.route("/reuniao/pauta", methods=["POST","GET"])
def pautas():
    #verifica se usuário está logado
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega usuario logado
        user = session["empresa"]
        #carrega banco de dados do usuario em questão
        dados = carregarReunioes(user)
        #lista aqueles a mais tempo não modificados
        dados = dados[["Data","Tema da reunião","Prioridade",'Concluido']]
        dados = dados.sort_values(['Data'])
        #lista daqueles com o prazo terminando em um dia
        prazo_vermelho = dados[dados['Prioridade']=="Urgente"]
        #lista daqueles com prazo entre 1 e 5 dias
        prazo_amarelo = dados[dados['Prioridade']=="Não Urgente"]

        #carrega html da rota /
        return render_template("listaPautas.html", tabela_modificado=dados.values.tolist(),tabela_prazo_vermelho= prazo_vermelho.values.tolist(),tabela_prazo_amarelo= prazo_amarelo.values.tolist())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))


















"""
Caminho referente aos Projetos TODO
"""



@app.route("/clientes/projetos", methods=["POST","GET"])
def projetos():
    if ("empresa" in session) and (session["nivel"]<2):
        #carrega usuário logado
        user = session["empresa"]
        #carrega banco de dados do usuario em questão
        projeto = carregarprojeto(user)
        if request.method == "POST":
            andamentos = carregarAndamentos(user)
            andamentos = andamentos[andamentos["DOC"]==int(request.form["doc"])].sort_values(["DATA"])
            return render_template("dados_andamento_cliente.html",tabela=clientes.values.tolist(),valores=andamentos.values.tolist())
        else:
            return render_template("dados_andamento_cliente.html",tabela=clientes.values.tolist())
    else:
        return redirect(url_for("login"))


# TODO : remover lead?
'''
@app.route("/leads/removercliente", methods=["POST","GET"])
def removerCliente():
    if "usuario" in session:
        user = session["usuario"]
        if request.method == "POST":
            clientes = carregarDados(user)
            DOC = int(request.form["DOC"])

            if DOC in clientes["DOC"].values:
                arq = clientes[clientes['DOC'] != DOC]
                arq = arq.sort_values(['JUSTIÇA','CLIENTE']).reset_index(drop=True).fillna('')
                arq.to_excel('clientes_'+user+'.xlsx')

                msg = "Cliente removido!!!"
                return render_template("remover.html", erro = msg)
            else:
                msg = "Sem cliente com este valor de DOC!"
                return render_template("remover.html", erro = msg)
        else:
            return render_template("remover.html")
    else:
        return redirect(url_for("login"))
'''
















"""
Feramentas globais
"""


def carregarCliente(usuario):
    try:
        try:
            return pd.read_excel(os.getcwd()+'/dados/clientes/clientes_luminus'+usuario+'.xlsx', sheet_name=0).drop(columns=["Unnamed: 0"]).reset_index(drop=True).fillna("")
        except:
            return pd.read_excel(os.getcwd()+'/dados/clientes/clientes_luminus'+usuario+'.xlsx', sheet_name=0).reset_index(drop=True).fillna("")
    except:
        return pd.DataFrame(columns=[colunas_coleta_leads])

def carregarReunioes(usuario):
    try:
        try:
            return pd.read_excel(os.getcwd()+'/dados/clientes/'+usuario+'/reunioes_'+usuario+'.xlsx', sheet_name=0).drop(columns=["Unnamed: 0"]).reset_index(drop=True).fillna("")
        except:
            return pd.read_excel(os.getcwd()+'/dados/clientes/'+usuario+'/reunioes_'+usuario+'.xlsx', sheet_name=0).reset_index(drop=True).fillna("")
    except:
        path = os.getcwd()+'/dados/clientes/'+usuario
        # Check whether the specified path exists or not
        isExist = os.path.exists(path)
        if not isExist:
            os.makedirs(path)

        return pd.DataFrame(columns=["Data","Participantes","Tema da reunião","Resumo da reunião","Anotações","Lista de decisões","Tarefas","Prioridade","Prazo","Concluido"])

def itemCliente():
    return pd.DataFrame(columns=[colunas_coleta_leads])

def adicionarValorCliente(local,clientes,nome,lugar,valor):
    index = clientes[clientes[local] == nome].index[0]
    for i in range(len(lugar)):
         clientes.at[index,lugar[i]] = valor[i].replace('\r', '')
    return clientes

def ConfirmarInt(numero):
    if numero.replace(" ","")!="":
        try:
            DOC = int(request.form[numero])
            return DOC
        except:
            msg = "Valor de "+numero+" Invalido"
            return render_template("adicionarClientes.html", erro = msg)
    else:
        msg = "Valor de "+numero+" Invalido"
        return render_template("adicionarClientes.html", erro = msg)

def ConfirmarTexto(numero):
    if str(numero.replace(" ",""))=="":
        msg = "Valor de "+numero+" Invalido"
        return render_template("adicionarClientes.html", erro = msg)
    else:
        return request.form[numero]

def ConfirmarData(numero):
    tratado = request.form[numero]
    tratado = tratado.replace(" ","")
    return tratado

def formatoData(valor):
    data = time.time() + 60*60*24*valor
    ano = str(time.localtime(data).tm_year)
    if (time.localtime(data).tm_mon<10):
       mes ="0" + str(time.localtime(data).tm_mon)
    else:
       mes = str(time.localtime(data).tm_mon)
    if (time.localtime(data).tm_mday<10):
       dia ="0" + str(time.localtime(data).tm_mday)
    else:
       dia = str(time.localtime(data).tm_mday)
    if (time.localtime(data).tm_hour<10):
       hora ="0" + str(time.localtime(data).tm_hour)
    else:
       hora = str(time.localtime(data).tm_hour)
    if (time.localtime(data).tm_min<10):
       minut ="0" + str(time.localtime(data).tm_min)
    else:
       minut = str(time.localtime(data).tm_min)
    if (time.localtime(data).tm_sec<10):
       seg ="0" + str(time.localtime(data).tm_sec)
    else:
       seg = str(time.localtime(data).tm_sec)
    return ano + "/" + mes + "/" + dia + "  " + hora + ":" + minut + ":" + seg


























































@app.route("/documentos")
def documentos():
    if "usuario" in session:
        return render_template("menu_documentos.html")
    else:
        return redirect(url_for("login"))

#rota de acesso /pdf/
#informações mostradas: html com as informações de todos os clientes em forma de tabela
@app.route('/documentos/pdf')
def pdf():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuario logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        #gera html do banco de dados do cliente que está em .xlmx
        clientes[clientes.columns[0:9]].to_html('templates/cliente_'+user+'.html')
        #codificação para teclado brasileiro
        options = {'encoding': "UTF-8"}
        #convertendo html para pdf
        pdfkit.from_file(os.getcwd()+'/templates/cliente_'+user+'.html', os.getcwd()+'/templates/cliente_'+user+'.pdf', options=options)
        #retorna html da tabela de clientes
        return send_file(os.getcwd()+"/templates/cliente_"+user+".pdf")
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))


#rota de acesso /exel
#Função: Mandar banco de dados completo dos clientes do usuário em formato .xlsx
@app.route("/documentos/exel")
def exel():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #retorna arquivo exel
        return send_file(os.getcwd()+'/clientes_'+user+'.xlsx')
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/documentos/procuracao", methods=["POST","GET"])
def procuracao():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]

        if request.method == "POST":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            cliente = clientes[clientes["DOC"]== int(request.form["doc"])]
            documento = Dc(os.getcwd()+"/documentos/modelo_formulário_procuracao_20_12_21.docx")
            alteracao = {'nome':cliente["NOME"][0].upper(),
                        'naturalidade':cliente["NATURALIDADE"][0],
                        'ocupacao':cliente["OCUPAÇÃO"][0],
                        'cpf':cliente["CPF"][0],
                        'rg':cliente["RG"][0],
                        'orgao':cliente["ORGÃO"][0],
                        'cidade':cliente["CIDADE"][0],
                        'endereco':cliente["ENDEREÇO"][0],
                        'email':cliente["EMAIL"][0],
                        'cell': cliente["CEL"][0],
                        'tipo_de_acao': request.form["tipo"],
                        'data': request.form["data"]
                        }
            documento.render(alteracao)
            documento.save(os.getcwd()+"/documentos/procuração_"+user+".docx")

            return send_file(os.getcwd()+'/documentos/procuração_'+user+'.docx')
        else:
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            #carrega a tela de pesquisa
            return render_template("procuracao.html", tabela=clientes.values.tolist())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#rota de acesso /pesquisar
#Função: Pesquisar por clientes a partir do nome ou trecho do nome
@app.route("/clientes/pesquisar", methods=["POST","GET"])
def pesquisar():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            #carrega termo que será pesquisado
            pesquisa = request.form["pesquisa"]
            #carrega clientes que coincidem com a pesquisa
            tabela_pesquisa = clientes[clientes['CLIENTE'].str.contains(pesquisa,case= False)]
            #carrega tela de pesquisa com a lista de resultados
            return render_template("pesquisa.html", tabela=tabela_pesquisa.values.tolist())
        else:
            #carrega a tela de pesquisa
            return render_template("pesquisa.html")
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#rota de acesso /modificar
#Função: Modifica dados de cliente
@app.route("/clientes/modificar", methods=["POST","GET"])
def modificar():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        tabela2 = pd.DataFrame(["JUSTIÇA COMUM","JUIZADO CÍVEL","JUIZADO CRIMINAL","JUSTIÇA FEDERAL","JUSTIÇA TRABALHISTA","ADMINISTRATIVO","INSS","OUTROS"])
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            #obtem valores preenchidos na página /modificar
            DOC = int(request.form["doc"])
            valor = request.form["colunas"]
            subs = request.form["subs"]
            justica = request.form["justica"]
            if(valor.upper()=="CLIENTE" and subs.replace(" ","")==""):
                #notifica mensagem de erro dos parametros inputados
                msg = "Nome do cliente invalido"
                return render_template("modificar.html", erro=msg,tabela=clientes.values.tolist(),colunas=clientes.columns[1:7],tabela2=tabela2.values.tolist())
            #verifica se cliente fornecido está na lista de clientes
            if DOC in clientes["DOC"].values:
                #verifica se a parâmetro é modificável
                if valor.upper() not in ['DOC','MODIFICADO','PRAZO']:
                    #verifica se parâmetro a ser modificável existe
                    if valor.upper() in clientes.columns:
                        #localiza cliente a ser modificado
                        index = clientes[clientes['DOC'] == DOC].index
                        #modifica valor no cliente
                        clientes.at[index,valor.upper()] = subs
                        clientes.at[index,"JUSTIÇA"] = justica
                        #salva a data de alteração
                        clientes.at[index,"MODIFICADO"] = formatoData(0)
                        #organiza a ordem
                        clientes = clientes.sort_values(['JUSTIÇA','CLIENTE']).reset_index(drop=True).fillna(' ')
                        #salva tabela modificada
                        try:
                            clientes.drop(columns=["Unnamed: 0"]).to_excel('clientes_'+user+'.xlsx')
                        except:
                            clientes.to_excel('clientes_'+user+'.xlsx')
                        #carrega pagina /modificar e notifica que alteração foi feita
                        msg = "CLIENTE MODIFICADO!!!!"
                        return render_template("modificar.html", erro=msg,tabela=clientes.values.tolist(),colunas=clientes.columns[1:7],tabela2=tabela2.values.tolist())
                    else:
                        #notifica mensagem de erro dos parametros inputados
                        msg = "Coluna inválida"
                        return render_template("modificar.html", erro=msg,tabela=clientes.values.tolist(),colunas=clientes.columns[1:7],tabela2=tabela2.values.tolist())
                else:
                    #notifica mensagem de erro dos parametros inputados
                    msg = "DOC, data de modificação e prazo não são alteráveis."
                    return render_template("modificar.html", erro=msg,tabela=clientes.values.tolist(),colunas=clientes.columns[1:7],tabela2=tabela2.values.tolist())
            else:
                #notifica mensagem de erro dos parametros inputados
                msg = "DOC não encontrado"
                return render_template("modificar.html", erro=msg,tabela=clientes.values.tolist(),colunas=clientes.columns[1:7],tabela2=tabela2.values.tolist())
        else:
            #carrega pagina /modificar
            return render_template("modificar.html",tabela=clientes.values.tolist(),colunas=clientes.columns[1:7],tabela2=tabela2.values.tolist())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/clientes/dados_cliente", methods=["POST","GET"])
def dados_cliente():
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        return render_template("dados_selecionar_cliente.html",tabela=clientes.values.tolist())
    else:
        return redirect(url_for("login"))

@app.route("/clientes/dados_cliente/dados", methods=["POST","GET"])
def dados():
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        if request.method == "POST" and "doc" in request.form:
            cliente = clientes[clientes["DOC"]== int(request.form["doc"])]
            return render_template("dados_cliente_selecionado.html",valor_doc=request.form["doc"],nome=cliente["NOME"][0],
                naturalidade=cliente["NATURALIDADE"][0],ocupacao=cliente["OCUPAÇÃO"][0],cpf=cliente["CPF"][0],rg=cliente["RG"][0],
                orgao=cliente["ORGÃO"][0],cidade=cliente["CIDADE"][0],endereco=cliente["ENDEREÇO"][0],email=cliente["EMAIL"][0],
                celular=cliente["CEL"][0])
        if request.method == "POST" and "pesquisa" in request.form:
            cliente = clientes[clientes["DOC"]== int(request.form["valor_doc_esc"])]
            cep = request.form["pesquisa"]
            if len(cep.replace("-","")) != 8:
                return
            correios = requests.get('https://viacep.com.br/ws/{}/json/'.format(cep.replace("-",""))).json()
            if 'erro' not in correios:
                cidade = correios['localidade']+"-"+correios['uf']
                endereco = correios['logradouro']+" "+correios['complemento']+", "+correios['bairro']+", "+correios['cep']
                return render_template("dados_cliente_selecionado.html",valor_doc=request.form["valor_doc_esc"],nome=cliente["NOME"][0],
                naturalidade=cliente["NATURALIDADE"][0],ocupacao=cliente["OCUPAÇÃO"][0],cpf=cliente["CPF"][0],rg=cliente["RG"][0],
                orgao=cliente["ORGÃO"][0],cidade=cidade,email=cliente["EMAIL"][0],
                celular=cliente["CEL"][0], endereco=endereco)

        if request.method == "POST" and "nome" in request.form:
            #if request.method == "POST" and request.form["pesquisa"]!="":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            cliente = clientes[clientes["DOC"] == int(request.form["valor_doc"])]
            #obtem valores preenchidos na página /modificar
            DOC = int(request.form["valor_doc"])
            nome = request.form["nome"]
            naturalidade = request.form["naturalidade"]
            ocupacao = request.form["ocupacao"]
            cpf = request.form["cpf"]
            rg = request.form["rg"]
            orgao = request.form["orgao"]
            cidade = request.form["cidade"]
            endereco = request.form["endereco"]
            email = request.form["email"]
            celular = request.form["celular"]
            #verifica se cliente fornecido está na lista de clientes
            if DOC in clientes["DOC"].values:
                #localiza cliente a ser modificado
                index = clientes[clientes['DOC'] == DOC].index
                #modifica valor no cliente
                clientes.at[index,"NOME"] = nome
                clientes.at[index,"NATURALIDADE"] = naturalidade
                clientes.at[index,"OCUPAÇÃO"] = ocupacao
                clientes.at[index,"CPF"] = cpf
                clientes.at[index,"RG"] = rg
                clientes.at[index,"ORGÃO"] = orgao
                clientes.at[index,"CIDADE"] = cidade
                clientes.at[index,"ENDEREÇO"] = endereco
                clientes.at[index,"EMAIL"] = email
                clientes.at[index,"CEL"] = celular
                #salva tabela modificada
                try:
                    clientes.drop(columns=["Unnamed: 0"]).to_excel('clientes_'+user+'.xlsx')
                except:
                    clientes.to_excel('clientes_'+user+'.xlsx')
                #carrega pagina /modificar e notifica que alteração foi feita
                msg = "CLIENTE ATUALIZADO!!!!"
                return render_template("dados_cliente_selecionado.html",valor_doc=request.form["valor_doc"],nome=cliente["NOME"][0],
                naturalidade=cliente["NATURALIDADE"][0],ocupacao=cliente["OCUPAÇÃO"][0],cpf=cliente["CPF"][0],rg=cliente["RG"][0],
                orgao=cliente["ORGÃO"][0],cidade=cliente["CIDADE"][0],endereco=cliente["ENDEREÇO"][0],email=cliente["EMAIL"][0],
                celular=cliente["CEL"][0],erro=msg)
            else:
                #notifica mensagem de erro dos parametros inputados
                msg = "DOC não encontrado"
                return render_template("dados_cliente_selecionado.html",valor_doc=request.form["valor_doc"],nome=cliente["NOME"][0],
                naturalidade=cliente["NATURALIDADE"][0],ocupacao=cliente["OCUPAÇÃO"][0],cpf=cliente["CPF"][0],rg=cliente["RG"][0],
                orgao=cliente["ORGÃO"][0],cidade=cliente["CIDADE"][0],endereco=cliente["ENDEREÇO"][0],email=cliente["EMAIL"][0],
                celular=cliente["CEL"][0],erro=msg)
    else:
        return redirect(url_for("login"))


@app.route("/clientes/andamentos/<DOC>", methods=["POST","GET"])
def andamentos_por_pesquisa(DOC):
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        clientes = clientes[clientes["DOC"]==int(DOC)]
        andamentos = carregarAndamentos(user)
        andamentos = andamentos[andamentos["DOC"]==int(DOC)].sort_values(["DATA"])
        return render_template("dados_andamento_cliente.html",tabela=clientes.values.tolist(),valores=andamentos.values.tolist())
    else:
        return redirect(url_for("login"))


@app.route("/andamento", methods=["POST","GET"])
def andamento():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            andamentos = carregarAndamentos(user)
            andamentos = andamentos[andamentos["DOC"].isin(clientes["DOC"].values)]
            #obtem valores preenchidos na página /andamento
            DOC = int(request.form["doc"])
            andamento = request.form["andamento"]
            #obtem a data no formato aaaa/mm/dd
            dia = int(request.form["dia"])
            if (dia<10):
               dia ="0" + str(dia)
            else:
               dia = str(dia)
            mes = int(request.form["mes"])
            if (mes<10):
               mes ="0" + str(mes)
            else:
               mes = str(mes)
            ano = str(request.form["ano"])
            data = ano+"/"+mes+"/"+dia
            #verifica se cliente fornecido está na lista de clientes
            if DOC in clientes["DOC"].values:
                #localiza cliente a ser atralado o prazo
                index = clientes[clientes['DOC'] == DOC].index.values[0]
                #modifica o prazo do cliente
                #print("AQUI!!!!!!!!!!!!!!!!!!!!!!!!!", index.values[0], DOC)
                clientes.at[index,"ANDAMENTO"] = data +"  "+andamento
                item = pd.DataFrame([[DOC,data,andamento]],columns=andamentos.columns)
                #salva a data de alteração
                clientes.at[index,"MODIFICADO"] = formatoData(0)
                #salva tabelas modificada
                try:
                    clientes.drop(columns=["Unnamed: 0"]).to_excel('clientes_'+user+'.xlsx')
                except:
                    clientes.to_excel('clientes_'+user+'.xlsx')
                try:
                    arq = pd.concat([andamentos,item], ignore_index=True).drop( columns=["Unnamed: 0"]).fillna('')
                except:
                    arq = pd.concat([andamentos,item], ignore_index=True).fillna('')
                arq.to_excel(os.getcwd()+'/dados/andamento_'+user+'.xlsx')
                #carrega pagina /prazo e notifica que alteração foi feita
                msg = "ANDAMENTO ADICIONADO!!!!"
                return render_template("andamento.html", erro=msg,tabela=clientes.values.tolist())
            else:
                #notifica mensagem de erro dos parametros inputados
                msg = "DOC não encontrado"
                return render_template("andamento.html", erro=msg,tabela=clientes.values.tolist())
        else:
            #carrega pagina /andamento
            return render_template("andamento.html",tabela=clientes.values.tolist())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

@app.route("/andamento/<doc>", methods=["POST","GET"])
def andamento_por_pesquisa(doc):
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        clientes = clientes[clientes["DOC"]==int(doc)]
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            andamentos = carregarAndamentos(user)
            andamentos = andamentos[andamentos["DOC"].isin(clientes["DOC"].values)]
            #obtem valores preenchidos na página /andamento
            DOC = int(request.form["doc"])
            andamento = request.form["andamento"]
            #obtem a data no formato aaaa/mm/dd
            dia = int(request.form["dia"])
            if (dia<10):
               dia ="0" + str(dia)
            else:
               dia = str(dia)
            mes = int(request.form["mes"])
            if (mes<10):
               mes ="0" + str(mes)
            else:
               mes = str(mes)
            ano = str(request.form["ano"])
            data = ano+"/"+mes+"/"+dia
            #verifica se cliente fornecido está na lista de clientes
            if DOC in clientes["DOC"].values:
                #localiza cliente a ser atralado o prazo
                index = clientes[clientes['DOC'] == DOC].index.values[0]
                #modifica o prazo do cliente
                #print("AQUI!!!!!!!!!!!!!!!!!!!!!!!!!", index.values[0], DOC)
                clientes.at[index,"ANDAMENTO"] = data +"  "+andamento
                item = pd.DataFrame([[DOC,data,andamento]],columns=andamentos.columns)
                #salva a data de alteração
                clientes.at[index,"MODIFICADO"] = formatoData(0)
                #salva tabelas modificada
                try:
                    clientes.drop(columns=["Unnamed: 0"]).to_excel('clientes_'+user+'.xlsx')
                except:
                    clientes.to_excel('clientes_'+user+'.xlsx')
                try:
                    arq = pd.concat([andamentos,item], ignore_index=True).drop( columns=["Unnamed: 0"]).fillna('')
                except:
                    arq = pd.concat([andamentos,item], ignore_index=True).fillna('')
                arq.to_excel(os.getcwd()+'/dados/andamento_'+user+'.xlsx')
                #carrega pagina /prazo e notifica que alteração foi feita
                msg = "ANDAMENTO ADICIONADO!!!!"
                return render_template("andamento.html", erro=msg,tabela=clientes.values.tolist())
            else:
                #notifica mensagem de erro dos parametros inputados
                msg = "DOC não encontrado"
                return render_template("andamento.html", erro=msg,tabela=clientes.values.tolist())
        else:
            #carrega pagina /andamento
            return render_template("andamento.html",tabela=clientes.values.tolist())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#rota de acesso /prazo
#Função: Atrelar um prazo ao cliente
@app.route("/prazo", methods=["POST","GET"])
def prazo():
    #verifica se usuário está logado
    if "usuario" in session:
        #carrega usuário logado
        user = session["usuario"]
        #carrega banco de dados do usuario em questão
        clientes = carregarDados(user)
        #verifica se foi feita requisição do tipo POST
        if request.method == "POST":
            #carrega banco de dados do usuario em questão
            clientes = carregarDados(user)
            #obtem valores preenchidos na página /prazo
            DOC = int(request.form["doc"])
            #obtem a data no formato aaaa/mm/dd
            dia = int(request.form["dia"])
            if (dia<10):
               dia ="0" + str(dia)
            else:
               dia = str(dia)
            mes = int(request.form["mes"])
            if (mes<10):
               mes ="0" + str(mes)
            else:
               mes = str(mes)
            ano = str(request.form["ano"])
            prazo = ano+"/"+mes+"/"+dia
            #verifica se cliente fornecido está na lista de clientes
            if DOC in clientes["DOC"].values:
                #localiza cliente a ser atralado o prazo
                index = clientes[clientes['DOC'] == DOC].index
                #modifica o prazo do cliente
                clientes.at[index,"PRAZO"] = prazo
                #salva a data de alteração
                clientes.at[index,"MODIFICADO"] = formatoData(0)
                #salva tabela modificada
                try:
                    clientes.drop(columns=["Unnamed: 0"]).to_excel('clientes_'+user+'.xlsx')
                except:
                    clientes.to_excel('clientes_'+user+'.xlsx')
                #carrega pagina /prazo e notifica que alteração foi feita
                msg = "PRAZO MODIFICADO!!!!"
                return render_template("prazo.html", erro=msg,tabela=clientes.values.tolist())
            else:
                #notifica mensagem de erro dos parametros inputados
                msg = "DOC não encontrado"
                return render_template("prazo.html", erro=msg,tabela=clientes.values.tolist())
        else:
            #carrega pagina /prazo
            return render_template("prazo.html",tabela=clientes.values.tolist())
    else:
        #redireciona para a tela de login
        return redirect(url_for("login"))

#rota de acesso /adicionar
#Função: Adiciona novo cliente
@app.route("/clientes/adicionar", methods=["POST","GET"])
def adicionar():
    #verifica se usuário está logado
    if "usuario" in session:
        user = session["usuario"]
        juizados = pd.DataFrame(["JUSTIÇA_COMUM","JUIZADO_CÍVEL","JUIZADO_CRIMINAL","JUSTIÇA_FEDERAL","JUSTIÇA_TRABALHISTA","ADMINISTRATIVO","INSS","OUTROS"])
        if request.method == "POST":
            clientes = carregarDados(user)
            try:
                DOC = int(request.form["doc"])
            except:
                msg = "Valor de DOC Invalido"
                return render_template("adicionar.html", erro = msg,tabela=juizados.values.tolist())
            cliente = request.form["cliente"][0:1].upper()
            cliente += request.form["cliente"][1:]
            if cliente.replace(" ","")=="" :
                msg = "Nome do cliente Invalido"
                return render_template("adicionar.html", erro = msg,tabela=juizados.values.tolist())
            vara = request.form["vara"]
            processo = request.form["processo"]
            distrib = request.form["distrib"]
            acao = request.form["acao"]
            nota = request.form["nota"]
            justica = request.form["justica"].replace("_"," ")
            andamento = ""
            modificado = formatoData(0)
            prazo = ""
            if len(clientes[clientes['DOC']==DOC].index) == 0:
                item = pd.DataFrame([[DOC,cliente,vara,processo,distrib,acao,nota,justica,andamento,modificado,prazo,"","","","","","","","","",""]],columns=clientes.columns)
                try:
                    arq = pd.concat([clientes,item], ignore_index=True).drop( columns=["Unnamed: 0"]).fillna('')
                except:
                    arq = pd.concat([clientes,item], ignore_index=True).fillna('')
                arq = arq.sort_values(["JUSTIÇA","CLIENTE"]).reset_index(drop=True).fillna('')
                arq.to_excel('clientes_'+user+'.xlsx')

                msg = "Cliente adicionado com sucesso!"
                return render_template("adicionar.html", erro = msg,tabela=juizados.values.tolist())
            else:
                msg = "Já existe cliente com este DOC"
                return render_template("adicionar.html", erro = msg,tabela=juizados.values.tolist())
        else:
            return render_template("adicionar.html",tabela=juizados.values.tolist())
    else:
        return redirect(url_for("login"))

@app.route("/clientes/remover", methods=["POST","GET"])
def remover():
    if "usuario" in session:
        user = session["usuario"]
        if request.method == "POST":
            clientes = carregarDados(user)
            DOC = int(request.form["DOC"])

            if DOC in clientes["DOC"].values:
                arq = clientes[clientes['DOC'] != DOC]
                arq = arq.sort_values(['JUSTIÇA','CLIENTE']).reset_index(drop=True).fillna('')
                arq.to_excel('clientes_'+user+'.xlsx')

                msg = "Cliente removido!!!"
                return render_template("remover.html", erro = msg)
            else:
                msg = "Sem cliente com este valor de DOC!"
                return render_template("remover.html", erro = msg)
        else:
            return render_template("remover.html")
    else:
        return redirect(url_for("login"))


@app.route("/clientes")
def clientes():
    if "usuario" in session:
        return render_template("menu_clientes.html")
    else:
        return redirect(url_for("login"))

def carregarDados(usuario):
    try:
        return pd.read_excel('clientes_'+usuario+'.xlsx', sheet_name=0).drop(columns=["Unnamed: 0"]).reset_index(drop=True).fillna("")
    except:
        return pd.read_excel('clientes_'+usuario+'.xlsx', sheet_name=0).reset_index(drop=True).fillna("")

def carregarAndamentos(usuario):
    try:
        return pd.read_excel(os.getcwd()+'/dados/andamento_'+usuario+'.xlsx', sheet_name=0).drop(columns=["Unnamed: 0"]).reset_index(drop=True).fillna("")
    except:
        return pd.read_excel(os.getcwd()+'/dados/andamento_'+usuario+'.xlsx', sheet_name=0).reset_index(drop=True).fillna("")



if __name__ == "__main__":
    app.run(debug=True, port=porta)