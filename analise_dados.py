import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

dataframe = pd.read_excel('C:\\Excel\\GLPI - Projetos.xlsx')

data = {
    'Status': ['Fechado', 'Andamento', 'Andamento', 'Fechado', 'Andamento', 'Fechado', 'Andamento', 'Andamento', 'Andamento', 'Fechado', 'Fechado', 'Fechado', 'Fechado', 'Fechado', 'Fechado', 'Fechado', 'Fechado', 'Andamento', 'Fechado', 'Andamento', 'Novo', 'Fechado', 'Novo', 'Fechado', 'Novo', 'Andamento', 'Aberto', 'Fechado', 'Fechado', 'Andamento', 'Andamento', 'Andamento', 'Fechado', 'Fechado', 'Andamento', 'Fechado', 'Fechado', 'Novo', 'Fechado', 'Fechado', 'Novo', 'Novo', 'Andamento'],
    'Nome': ['Impressão de Etiquetas', 'Sistema Laboratório', 'Consulta personalizada Composição Teares', 'Implantar GLPI', 'Interface para armazenamento de senhas', 'Relatório personalizado de Crédito', 'App Vendas', 'Machine Learnirg', 'Interface para unir, dividir e converter pdf', 'Consulta personalizada Saldo Inicial/Final', 'Desenvolver Query Produção', 'Desenvovler Query Vendas', 'Desenvolver Query Estoque', 'Desenvolver Painel Analise de Crédito', 'Desenvolver Painel Estoque x Mapa de Produção', 'Desenvolver Query Apontamento de Produção', 'Desenvolver Painel Apontamento de Produtividade', 'Integração Bancaria Caixa', 'Integração Bancaria Daycoval', 'Integração Bancaria Sicredi', 'Sistema de Localização de Estoque', 'Hospedagem Site Stik', 'Organização Hack PCP', 'Automação Painel Embalagem', 'Atualizar Versão Top Manager', 'Desenvolver Painel PCP', 'Desenvolver Query Saldo Pedido Tabela', 'Consulta personalizada Consumo Recompra Cliente', 'Consulta personalizada Consulta Media Estoque', 'Desenvolver painel de projetos em andamentos', 'Implementações de telas no site Stik', 'Criação de DashBoard para contabilizar chamado RCN & BS', 'Criação de programa de melhoria a rotina Administrativa (PDF) V1', 'Criação de programa de melhoria ds rotina Administrativa (KEYLOGGER) V1', 'Consulta personalizada Mapeamento Falha Tipo B', 'Consulta personalizada Saving 2.0', 'Consulta personalizada Sugestão de Compra', 'Desenvolver Etiquetas Novos Tamanhos', 'Desenvolver Query Pedidos Incluidos ', 'Desenvolver Painel Capacidade Produtiva', 'Otimizar processo de reserva tela V1', 'Desenvolver tela inclusão metragem revisão', 'Monitoramento Backup File Server'],
    'Data planejada para começo': ['17/07/2023', '01/05/2023', '20/07/2023', '17/07/2023', '01/09/2023', '28/08/2023', '15/08/2023', '30/09/2023', '01/08/2023', '01/07/2023', '20/04/2023', '01/06/2023', '01/05/2023', '31/08/2023', '31/07/2023', '31/01/2023', '06/04/2023', '15/06/2023', '30/06/2023', '02/08/2023', '18/09/2023', '10/09/2023', '01/12/2023', '10/05/2023', '19/09/2023', '25/09/2023', '25/09/2023', '06/10/2023', '25/09/2023', '19/10/2023', '20/10/2023', '17/10/2023', '12/10/2023', '10/08/2023', '25/09/2023', '06/10/2023', '11/10/2023', '06/10/2023', '01/11/2023', '15/10/2023','08/11/2023', '10/11/2023'],
    'Data planejada para fim': ['31/07/2023', '30/09/2023', '30/09/2023', '20/09/2023', '30/09/2023', '28/09/2023', '31/10/2023', '15/11/2023', '30/09/2023', '31/07/2023', '16/08/2023', '31/08/2023', '31/07/2023', '15/09/2023', '16/08/2023', '01/04/2023', '30/04/2023', '01/08/2023', '31/07/2023', '30/09/2023', '30/09/2023', '20/09/2023', '15/12/2023', '13/05/2023', '07/10/2023', '07/10/2023', '25/10/2023', '27/10/2023', '24/10/2023', '27/10/2023', '27/10/2023', '25/10/2023', '02/10/2023', '29/09/2023', '17/10/2023', '08/11/2023', '08/12/2023', '03/11/2023', '08/11/2023', '20/11/2023', '08/12/2023', '10/12/2023'],
}   

df = pd.DataFrame(dataframe)

plt.figure(figsize=(10, 6))
sns.set(style='whitegrid')
sns.countplot(data=df, x='Status')

ax = sns.countplot(data=df, x='Status') 

for p in ax.patches:
    ax.annotate(f'{p.get_height()}', (p.get_x() + p.get_width() / 2., p.get_height()),
    ha='center', va='center', fontsize=8, color='black', xytext=(0, 10),
    textcoords='offset points')
plt.title('Contagem de Projetos por Status', fontsize=14)
plt.xlabel('Status', fontsize=12)
plt.ylabel('Contagem', fontsize=12)
plt.xticks(rotation=360, ha='right')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()

plt.show()

df['Data planejada para começo'] = pd.to_datetime(df['Data planejada para começo'], format='%d/%m/%Y', errors='coerce')
df['Data planejada para fim'] = pd.to_datetime(df['Data planejada para fim'], format='%d/%m/%Y', errors='coerce')
df['Duração planejada (dias)'] = (df['Data planejada para fim'] - df['Data planejada para começo']).dt.days

projetos_mais_demorados = df.sort_values('Duração planejada (dias)', ascending=False)['Nome'].drop_duplicates()

projetos_top = df[df['Nome'].isin(projetos_mais_demorados.head(45))]
projetos_top = projetos_top[['Nome', 'Duração planejada (dias)','Status']].drop_duplicates()

print(projetos_top)

