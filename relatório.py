import pandas as pd
from tkinter import filedialog
from tkinter import Tk
import locale
from datetime import datetime
from fpdf import FPDF

# Configurar o locale para o formato de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Função para ler o arquivo XLS e processar os dados
def ler_e_processar_xls(arquivo_xls):
    # Ler o arquivo XLS
    df = pd.read_excel(arquivo_xls)
    
    # Garantir que os valores monetários sejam floats
    df['valor'] = df['valor'].apply(lambda x: float(str(x).replace(',', '.')))
    
    # Criar um dicionário para facilitar a manipulação dos dados
    dados = df.set_index('Identificação')['valor'].to_dict()
    
    return dados

# Função para calcular os valores solicitados
def calcular_valores(dados):
    # Constantes
    salario_minimo = 1320.00
    adicional_noturno = 0.20
    taxa_insalubridade = 0.15
    
    # Cálculos
    quantidade_unidades = dados.get('Quantidade de unidades no total', 1)
    valor_total_agua = dados.get('Conta de agua', 0)
    valor_total_luz = dados.get('Conta de luz', 0)
    valor_agua_por_unidade = valor_total_agua / quantidade_unidades
    valor_luz_por_unidade = valor_total_luz / quantidade_unidades
    valor_taxa_condominial = dados.get('Taxa condominial', 0)
    
    # Total de gastos com funcionários
    funcionarios_total = dados.get('Quantidade de funcionários', 0)
    funcionarios_limpeza = dados.get('Funcionários da limpeza', 0)
    funcionarios_noturno = dados.get('Funcionários do período noturno', 0)
    
    salario_funcionarios = (funcionarios_total - funcionarios_limpeza - funcionarios_noturno) * salario_minimo
    salario_limpeza = funcionarios_limpeza * salario_minimo * (1 + taxa_insalubridade)
    salario_noturno = funcionarios_noturno * salario_minimo * (1 + adicional_noturno)
    
    total_gasto_funcionarios = salario_funcionarios + salario_limpeza + salario_noturno
    
    # Salário individual com adicional noturno e com taxa de insalubridade
    salario_individual_adicional_noturno = salario_minimo * (1 + adicional_noturno)
    salario_individual_insalubridade = salario_minimo * (1 + taxa_insalubridade)
    
    # Valor a pagar por unidade
    valor_pagar_por_unidade = (total_gasto_funcionarios + valor_total_agua + valor_total_luz + valor_taxa_condominial) / quantidade_unidades
    
    # Valor total do condomínio
    valor_total_condominio = valor_pagar_por_unidade + valor_taxa_condominial
    
    return {
        'valor_total_agua': valor_total_agua,
        'valor_agua_por_unidade': valor_agua_por_unidade,
        'valor_total_luz': valor_total_luz,
        'valor_luz_por_unidade': valor_luz_por_unidade,
        'valor_taxa_condominial': valor_taxa_condominial,
        'total_gasto_funcionarios': total_gasto_funcionarios,
        'salario_noturno': salario_noturno,
        'salario_funcionarios': salario_funcionarios,
        'salario_individual_adicional_noturno': salario_individual_adicional_noturno,
        'salario_individual_insalubridade': salario_individual_insalubridade,
        'valor_pagar_por_unidade': valor_pagar_por_unidade,
        'valor_total_condominio': valor_total_condominio,
        'total_despesas_mensais': valor_total_agua + valor_total_luz + total_gasto_funcionarios + valor_taxa_condominial
    }

# Função para formatar valores em reais
def formatar_reais(valor):
    return locale.currency(valor, grouping=True)

# Função para gerar o relatório em um arquivo PDF
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Relatório de Gestão do Condomínio', 0, 1, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(5)

    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, body)
        self.ln()

def gerar_relatorio(dados_calculados, caminho_relatorio):
    # Data atual
    data_atual = datetime.now().strftime('%B/%Y')
    
    pdf = PDF()
    pdf.add_page()
    
    # Título e Introdução
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Relatório de Gestão do Condomínio', 0, 1, 'C')
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 10, f'Condomínio: "14 de Setembro"', 0, 1, 'C')
    pdf.cell(0, 10, f'Período: {data_atual}', 0, 1, 'C')
    pdf.ln(10)
    
    # Resumo Executivo
    pdf.chapter_title('Resumo Executivo')
    pdf.chapter_body(
        'Este relatório apresenta um resumo das contas e despesas do condomínio "14 de Setembro" '
        f'para o período mencionado. Inclui uma análise detalhada dos gastos com água, luz, taxa condominial e funcionários.'
    )
    
    # Contas Mensais
    pdf.chapter_title('Contas Mensais')
    pdf.chapter_title('Água')
    pdf.chapter_body(
        f"Valor total da conta de água: {formatar_reais(dados_calculados['valor_total_agua'])}\n"
        f"Valor da conta de água por unidade: {formatar_reais(dados_calculados['valor_agua_por_unidade'])}"
    )
    
    pdf.chapter_title('Luz')
    pdf.chapter_body(
        f"Valor total da conta de luz: {formatar_reais(dados_calculados['valor_total_luz'])}\n"
        f"Valor da conta de luz por unidade: {formatar_reais(dados_calculados['valor_luz_por_unidade'])}"
    )
    
    total_gasto_contas = dados_calculados['valor_total_agua'] + dados_calculados['valor_total_luz']
    pdf.chapter_title('Total de Gasto com Contas')
    pdf.chapter_body(
        f"Total: {formatar_reais(total_gasto_contas)}\n"
        f"Valor a pagar por unidade: {formatar_reais(dados_calculados['valor_pagar_por_unidade'])}"
    )
    
    # Taxa Condominial
    pdf.chapter_title('Taxa Condominial')
    pdf.chapter_body(
        f"Valor da taxa condominial: {formatar_reais(dados_calculados['valor_taxa_condominial'])}\n"
        f"Valor total do condomínio por unidade (incluindo contas): {formatar_reais(dados_calculados['valor_total_condominio'])}"
    )
    
    # Gastos com Funcionários
    pdf.chapter_title('Gastos com Funcionários')
    
    # Adicionar texto explicativo
    texto_explicativo = (
        "Os gastos com funcionários são calculados com base nos seguintes valores:\n"
        "- Salário mínimo vigente: R$ 1320,00\n"
        "- Taxa de insalubridade: 15%\n"
        "- Adicional noturno: 20%\n\n"
        "Cálculo Detalhado:\n"
        f"Salário individual com adicional noturno: {formatar_reais(dados_calculados['salario_individual_adicional_noturno'])}\n"
        f"Salário individual com taxa de insalubridade: {formatar_reais(dados_calculados['salario_individual_insalubridade'])}\n"
        f"Total de gasto com funcionários: {formatar_reais(dados_calculados['total_gasto_funcionarios'])}\n"
        f"Valor a pagar por unidade: {formatar_reais(dados_calculados['valor_pagar_por_unidade'])}"
    )
    
    pdf.chapter_body(texto_explicativo)
    
    # Conclusão
    pdf.chapter_title('Conclusão')
    pdf.chapter_body(
        f'O total das despesas mensais do condomínio "14 de Setembro" para o período mencionado é de {formatar_reais(dados_calculados["total_despesas_mensais"])}'
        f', o que resulta em um custo mensal por unidade de {formatar_reais(dados_calculados["valor_total_condominio"])}.'
        ' É importante que todos os moradores estejam cientes dessas despesas para garantir uma gestão financeira eficiente e a manutenção dos serviços e facilidades oferecidos pelo condomínio.'
    )
    
    pdf.output(caminho_relatorio)

# Função principal para criar a interface e executar as funções
def main():
    # Abrir janela de seleção de arquivo
    root = Tk()
    root.withdraw()  # Esconder a janela principal
    arquivo_xls = filedialog.askopenfilename(title="Selecione o arquivo XLS")
    
    if arquivo_xls:
        # Executar as funções
        dados = ler_e_processar_xls(arquivo_xls)
        dados_calculados = calcular_valores(dados)
        
        caminho_relatorio = 'relatorio_condominio.pdf'
        gerar_relatorio(dados_calculados, caminho_relatorio)
        
        print("Relatório gerado com sucesso!")
    else:
        print("Nenhum arquivo selecionado.")

if __name__ == "__main__":
    main()
