from docxtpl import DocxTemplate, RichText
import pandas as pd
import unidecode
import tkinter as tk
import os
from pypodio2 import api


project_dir = os.path.dirname(__file__)
termo = project_dir + r'\\Termos\\'

doc = DocxTemplate("termo_template_pc.docx")
#
# #
# # doc = DocxTemplate("template-termo.docx")
# # rt = RichText()
# # rt.add(person, bold=True, font='Calibri', size=20)
# # rt_1 = RichText()
# # rt_1.add(location, bold=True, font='Calibri', size=20)
# # context = {'workplan_number': '102', 'person': rt, 'location': rt_1}
# # doc.render(context)
# # doc.save('YT_template77.docx')
#
#
def retorna_patrimonio_por_departamento(departamento):
    """
    Retorna um dicionário de listas com os números de patrimônio das duas respetivas colunas
    :param
    :return
    """
    df_inventario = pd.read_csv('inventario.csv', dtype=str)
    df_filtro_departamento = df_inventario.loc[df_inventario['Departamento'].isin([departamento])].fillna('#')
    dicionario_patrimonios_desktops = df_filtro_departamento[['Patrimônio Desktop ou notebook']].to_dict('list')
    dicionario_patrimonios_monitores = df_filtro_departamento[['Patrimônio Monitor']].to_dict('list')
    return dicionario_patrimonios_desktops, dicionario_patrimonios_monitores


def retorna_dados_funcionario(nome):
    df_funcionarios = pd.read_csv('base_funcionarios.csv')
    df_filtro = df_funcionarios.loc[df_funcionarios['NOMEFUNCIONARIO'] == nome]
    dicionario_dados_funcionarios = df_filtro[['MATRICULA', 'NOMEFUNCIONARIO', 'CPF', 'FUNCAO']].to_dict('list')
    return dicionario_dados_funcionarios

client = api.OAuthClient(
    'tablets-novo',
    'SECRET',
    'USERr',
    'PASSWORD',
)


def filtro_app_id(departamento):
    app = client.Application.find(22771792)
    filtro = client.Item.filter(22771792, {'filters': {'title':f'{departamento}'}})
    chave = []
    valor = []
    api_tuple = []
    for key in filtro['items']:
        for i in key['fields']:
            for a, b in i.items():
                chave.append(a)
                valor.append(b)
    api_tuple = list(zip(chave, valor))
    df = pd.DataFrame(api_tuple, columns=['Campo', 'Valor'], dtype=str)
    series_num_patrimonios_desktop = df.loc[df.shift(1)['Valor'] == 'Patrimônio Desktop ou notebook']['Valor'].to_frame()
    dicio_patrimonios_desktop = series_num_patrimonios_desktop.to_dict('records')
    series_num_patrimonios_monitor = df.loc[df.shift(1)['Valor'] == 'Patrimônio Monitor']['Valor'].to_frame()
    dicio_patrimonios_monitor = series_num_patrimonios_monitor.to_dict('records')

# Bloco de extração dos dígitos desktop
    c = 0
    dict_desktop_notebook_patrimonios = {}
    while c < len(dicio_patrimonios_desktop):
        temp_dict = dicio_patrimonios_desktop[c]
        temp_list =[]
        for value in temp_dict.values():
            for digit in value:
                if digit.isdigit():
                    temp_list.append(digit)
            temp_list = ''.join(temp_list)
            dict_desktop_notebook_patrimonios[f'pat_desk_{c+1}'] = temp_list
        c += 1

    # Bloco de extração dos dígitos monitores
    c = 0
    dict_monitor_patrimonios = {}
    while c < len(dicio_patrimonios_monitor):
        temp_dict = dicio_patrimonios_monitor[c]
        temp_list =[]
        for value in temp_dict.values():
            for digit in value:
                if digit.isdigit():
                    temp_list.append(digit)
            temp_list = ''.join(temp_list)
            dict_monitor_patrimonios[f'pat_mon_{c+1}'] = temp_list
        c += 1

    return dict_desktop_notebook_patrimonios, dict_monitor_patrimonios

def gera_context_patrimonios(desktops, monitores):
    """
    Desempacota lista de patrimônios, retorna dicionário para renderização
    :param desktops: lista de desktops do dataframe
    :param monitores: lista de monitores do dataframe
    :return: dicionário no formato padrão para renderização no docx
    """
    # lista_desktops_patrimonios_valor = [item for _, lista in desktops.items() for item in lista]
    # lista_desktops_chave = [f'pat_desk_{n + 1}' for n in range(len(lista_desktops_patrimonios_valor))]
    # context_desktops = dict(zip(lista_desktops_chave, lista_desktops_patrimonios_valor))
    #
    # lista_monitores_patrimonios_valor = [item for _, lista in monitores.items() for item in lista]
    # lista_monitores_chave = [f'pat_mon_{n + 1}' for n in range(len(lista_monitores_patrimonios_valor))]
    # context_monitores = dict(zip(lista_monitores_chave, lista_monitores_patrimonios_valor))
    context_desktops = desktops
    context_monitores = monitores

    dicionario_context_patrimonios = {**context_desktops, **context_monitores}
    return dicionario_context_patrimonios


def gera_context_funcionario(nome):
    lista_dados_funcionarios_valores = [item for _, lista in nome.items() for item in lista]
    lista_dados_funcionarios_chaves = ['num_matricula', 'nome_colab', 'num_cpf', 'cargo_func']
    context_funcionarios = dict(zip(lista_dados_funcionarios_chaves, lista_dados_funcionarios_valores))
    return context_funcionarios


def mescla_context(patrimonio, funcionario):
    context_final = {**patrimonio, **funcionario}
    return context_final


def renderiza_termo(context, dep):
    doc.render(context)
    doc.save(f'Termo {dep}.docx')
    return


def main(n, d):
    input_departamento = 'LOJA VILARINHO'
    desktops, monitores = filtro_app_id(d)
    dicio_context_patrimonios = gera_context_patrimonios(desktops, monitores)
    dep = d
    funcionario = retorna_dados_funcionario(n)
    dicio_context_funcionario = gera_context_funcionario(funcionario)
    context_final = mescla_context(dicio_context_patrimonios, dicio_context_funcionario)
    renderiza_termo(context_final, dep)

def get_input():
    n = str(nome.get()).upper()
    n = unidecode.unidecode(n)
    d = str(departamento.get()).upper()
    d = unidecode.unidecode(d)
    return main(n, d)


app = tk.Tk()
app.title('Gerador de Termo de Uso Desktop Notebooks')
app.geometry('400x200')
app.configure(background='#dde')
tk.Label(app, text='Nome Colaborador', background='#dde', foreground='#009', anchor='nw').place(x=150, y=10, width=150,
                                                                                    height=20)
nome = tk.Entry(app)
nome.place(x=50, y=35, width=300, height=20)

tk.Label(app, text='Departamento', background='#dde', foreground='#009', anchor='nw').place(x=155, y=60, width=150,
                                                                                          height=20)
departamento = tk.Entry(app)

departamento.place(x=50, y=80, width=300, height=20)
tk.Button(app, text='Gerar Termo', command=get_input).place(x=150, y=150, width=100)
app.mainloop()


