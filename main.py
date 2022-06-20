from docxtpl import DocxTemplate, RichText
import pandas as pd

doc = DocxTemplate("termo_template_pc.docx")


# person = 'José das Couves'.upper()
# location = 'São José das Couves'.upper()
#
# doc = DocxTemplate("template-termo.docx")
# rt = RichText()
# rt.add(person, bold=True, font='Calibri', size=20)
# rt_1 = RichText()
# rt_1.add(location, bold=True, font='Calibri', size=20)
# context = {'workplan_number': '102', 'person': rt, 'location': rt_1}
# doc.render(context)
# doc.save('YT_template77.docx')


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


def gera_context_patrimonios(desktops, monitores):
    """
    Desempacota lista de patrimônios, retorna dicionário para renderização
    :param desktops: lista de desktops do dataframe
    :param monitores: lista de monitores do dataframe
    :return: dicionário no formato padrão para renderização no docx
    """
    lista_desktops_patrimonios_valor = [item for _, lista in desktops.items() for item in lista]
    lista_desktops_chave = [f'pat_desk_{n + 1}' for n in range(len(lista_desktops_patrimonios_valor))]
    context_desktops = dict(zip(lista_desktops_chave, lista_desktops_patrimonios_valor))

    lista_monitores_patrimonios_valor = [item for _, lista in monitores.items() for item in lista]
    lista_monitores_chave = [f'pat_mon_{n + 1}' for n in range(len(lista_monitores_patrimonios_valor))]
    context_monitores = dict(zip(lista_monitores_chave, lista_monitores_patrimonios_valor))

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


def renderiza_termo(context):
    doc.render(context)
    doc.save('Teste novissímo.docx')
    return


def main():
    input_departamento = 'LOJA VILARINHO'
    desktops, monitores = retorna_patrimonio_por_departamento(input_departamento)
    dicio_context_patrimonios = gera_context_patrimonios(desktops, monitores)
    input_colaborador = ''
    funcionario = retorna_dados_funcionario(input_colaborador)
    dicio_context_funcionario = gera_context_funcionario(funcionario)
    context_final = mescla_context(dicio_context_patrimonios, dicio_context_funcionario)
    renderiza_termo(context_final)


if __name__ == '__main__':
    main()
