#!/usr/bin/env python3
"""
Script para criar arquivo Excel de exemplo para o RPA de Notas Fiscais
Execute este script para gerar o arquivo notas_fiscais.xlsx
"""

import pandas as pd
from datetime import datetime

def criar_excel_exemplo():
    """Cria arquivo Excel de exemplo com os dados das notas fiscais"""
    
    # Dados de exemplo baseados na tabela HTML fornecida
    dados = [
        {
            'Data': '16/09/25',
            'Nome_Cliente': 'Ana Claudia',
            'CPF': '28471111829',
            'Nome_Pet': 'Lili',
            'Tipo_Servico': 'Consulta Veterinária',
            'Valor': 350.00,
            'Endereco': 'R CLaudemeses dos Santos, 164',
            'Cidade': 'Valinhos'
        },
        {
            'Data': '16/09/25',
            'Nome_Cliente': 'Paula Codina',
            'CPF': '24924940801',
            'Nome_Pet': 'Boy',
            'Tipo_Servico': 'Banho e Tosa',
            'Valor': 300.00,
            'Endereco': 'R dez, 1335',
            'Cidade': 'Campinas'
        },
        {
            'Data': '17/09/25',
            'Nome_Cliente': 'Edvandro Trindade',
            'CPF': '13777593850',
            'Nome_Pet': 'Floquinho',
            'Tipo_Servico': 'Consulta Clínica',
            'Valor': 350.00,
            'Endereco': 'R 24 de Junho, 338',
            'Cidade': 'Capivari'
        },
        {
            'Data': '18/09/25',
            'Nome_Cliente': 'Maria Azevedo',
            'CPF': '22489089897',
            'Nome_Pet': 'Bruna',
            'Tipo_Servico': 'Medicamento Animal',
            'Valor': 300.00,
            'Endereco': 'R Helio Martins, 143',
            'Cidade': 'Sumare'
        },
        {
            'Data': '19/09/25',
            'Nome_Cliente': 'Marcia Prezoto',
            'CPF': '09057375800',
            'Nome_Pet': 'Pepe',
            'Tipo_Servico': 'Medicamento Animal',
            'Valor': 300.00,
            'Endereco': 'R Adolfo de Oliveira, 100',
            'Cidade': 'Sumare'
        },
        {
            'Data': '19/05/25',
            'Nome_Cliente': 'Rosangela Magalhes',
            'CPF': '05187618865',
            'Nome_Pet': 'Branquinha',
            'Tipo_Servico': 'Medicamento Animal',
            'Valor': 330.00,
            'Endereco': 'R 8, 181',
            'Cidade': 'Sumare'
        },
        # Dados adicionais para exemplo mais completo
        {
            'Data': '20/09/25',
            'Nome_Cliente': 'João Silva',
            'CPF': '12345678901',
            'Nome_Pet': 'Rex',
            'Tipo_Servico': 'Vacina',
            'Valor': 150.00,
            'Endereco': 'R das Flores, 456',
            'Cidade': 'Indaiatuba'
        },
        {
            'Data': '21/09/25',
            'Nome_Cliente': 'Maria Santos',
            'CPF': '98765432100',
            'Nome_Pet': 'Luna',
            'Tipo_Servico': 'Cirurgia',
            'Valor': 800.00,
            'Endereco': 'Av Principal, 789',
            'Cidade': 'Campinas'
        },
        {
            'Data': '22/09/25',
            'Nome_Cliente': 'Carlos Oliveira',
            'CPF': '11122233344',
            'Nome_Pet': 'Miau',
            'Tipo_Servico': 'Exame',
            'Valor': 120.00,
            'Endereco': 'R do Centro, 25',
            'Cidade': 'Valinhos'
        },
        {
            'Data': '23/09/25',
            'Nome_Cliente': 'Ana Paula Costa',
            'CPF': '55566677788',
            'Nome_Pet': 'Thor',
            'Tipo_Servico': 'Consulta de Rotina',
            'Valor': 250.00,
            'Endereco': 'R Nova Esperança, 88',
            'Cidade': 'Sumare'
        }
    ]
    
    # Criar DataFrame
    df = pd.DataFrame(dados)
    
    # Configurar o ExcelWriter com formatação
    with pd.ExcelWriter('notas_fiscais.xlsx', engine='openpyxl') as writer:
        # Escrever dados na planilha
        df.to_excel(writer, sheet_name='Dados', index=False)
        
        # Obter a worksheet para formatação
        worksheet = writer.sheets['Dados']
        
        # Ajustar largura das colunas
        column_widths = {
            'A': 12,  # Data
            'B': 20,  # Nome_Cliente
            'C': 15,  # CPF
            'D': 15,  # Nome_Pet
            'E': 20,  # Tipo_Servico
            'F': 10,  # Valor
            'G': 35,  # Endereco
            'H': 15   # Cidade
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Formatar cabeçalho
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)
            cell.fill = cell.fill.copy(fgColor="DDDDDD")
    
    print("✅ Arquivo 'notas_fiscais.xlsx' criado com sucesso!")
    print(f"📊 Total de registros: {len(dados)}")
    print("\n📋 Estrutura criada:")
    print(df.head())
    print(f"\n📍 Cidades incluídas: {', '.join(df['Cidade'].unique())}")
    print(f"💰 Valores: R$ {df['Valor'].min():.2f} - R$ {df['Valor'].max():.2f}")

def validar_dados():
    """Valida o arquivo criado"""
    try:
        df = pd.read_excel('notas_fiscais.xlsx')
        
        print("\n🔍 Validação dos dados:")
        
        # Verificar colunas obrigatórias
        colunas_obrigatorias = ['Data', 'Nome_Cliente', 'CPF', 'Nome_Pet', 'Tipo_Servico', 'Valor', 'Endereco', 'Cidade']
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df.columns]
        
        if colunas_faltando:
            print(f"❌ Colunas faltando: {colunas_faltando}")
        else:
            print("✅ Todas as colunas obrigatórias estão presentes")
        
        # Verificar CPFs
        cpfs_invalidos = df[df['CPF'].astype(str).str.len() != 11]
        if len(cpfs_invalidos) > 0:
            print(f"⚠️  CPFs com formato incorreto: {len(cpfs_invalidos)}")
        else:
            print("✅ Todos os CPFs têm 11 dígitos")
        
        # Verificar valores
        valores_invalidos = df[df['Valor'] <= 0]
        if len(valores_invalidos) > 0:
            print(f"⚠️  Valores inválidos: {len(valores_invalidos)}")
        else:
            print("✅ Todos os valores são positivos")
        
        # Mostrar estatísticas
        print(f"\n📈 Estatísticas:")
        print(f"   • Total de registros: {len(df)}")
        print(f"   • Valor total: R$ {df['Valor'].sum():.2f}")
        print(f"   • Valor médio: R$ {df['Valor'].mean():.2f}")
        print(f"   • Cidades únicas: {df['Cidade'].nunique()}")
        print(f"   • Clientes únicos: {df['Nome_Cliente'].nunique()}")
        
        return True
        
    except FileNotFoundError:
        print("❌ Arquivo não encontrado. Execute primeiro a criação do arquivo.")
        return False
    except Exception as e:
        print(f"❌ Erro na validação: {str(e)}")
        return False

def criar_template_vazio():
    """Cria um template vazio para preenchimento manual"""
    
    # Criar DataFrame vazio com apenas os cabeçalhos
    template = pd.DataFrame(columns=[
        'Data', 'Nome_Cliente', 'CPF', 'Nome_Pet', 
        'Tipo_Servico', 'Valor', 'Endereco', 'Cidade'
    ])
    
    # Adicionar algumas linhas de exemplo comentadas
    exemplos = [
        "# Exemplo: 16/09/25, Ana Claudia, 28471111829, Lili, Consulta, 350.00, R das Flores 123, Valinhos",
        "# Data no formato DD/MM/AA",
        "# CPF com 11 dígitos, sem pontos ou hífens", 
        "# Valor com ponto decimal (350.00)",
        "# Endereço: R Nome da Rua, Número"
    ]
    
    # Salvar template
    with pd.ExcelWriter('template_notas_fiscais.xlsx', engine='openpyxl') as writer:
        template.to_excel(writer, sheet_name='Template', index=False)
        
        worksheet = writer.sheets['Template']
        
        # Ajustar larguras
        column_widths = {'A': 12, 'B': 20, 'C': 15, 'D': 15, 'E': 20, 'F': 10, 'G': 35, 'H': 15}
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Formatar cabeçalho
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)
            cell.fill = cell.fill.copy(fgColor="DDDDDD")
    
    print("✅ Template vazio criado: 'template_notas_fiscais.xlsx'")

if __name__ == "__main__":
    print("🏥 Gerador de Excel para RPA de Notas Fiscais Veterinárias")
    print("=" * 55)
    
    # Menu de opções
    print("\nEscolha uma opção:")
    print("1. Criar arquivo com dados de exemplo")
    print("2. Criar template vazio para preenchimento manual")
    print("3. Validar arquivo existente")
    print("4. Criar ambos (exemplo + template)")
    
    opcao = input("\nDigite sua opção (1-4): ").strip()
    
    if opcao == "1":
        criar_excel_exemplo()
        validar_dados()
    elif opcao == "2":
        criar_template_vazio()
    elif opcao == "3":
        validar_dados()
    elif opcao == "4":
        criar_excel_exemplo()
        criar_template_vazio()
        validar_dados()
    else:
        print("❌ Opção inválida. Execute novamente.")
    
    print("\n" + "=" * 55)
    print("Arquivos prontos para usar com o RPA!")