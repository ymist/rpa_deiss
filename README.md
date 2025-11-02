# 🤖 RPA para Preenchimento Automático de Formulários Web

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/Selenium-4.15+-green.svg)](https://www.selenium.dev/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

> **Sistema automatizado de preenchimento de formulários web que reduziu processo manual de 2 dias para execução automatizada em minutos.**

## 📋 Sobre o Projeto

Este RPA (Robotic Process Automation) foi desenvolvido para automatizar o preenchimento de notas fiscais eletrônicas para o sistema DEISS, eliminando trabalho manual repetitivo e reduzindo drasticamente o tempo de processamento.

### 🎯 Problema Resolvido

-   **Antes:** Processo manual de preenchimento levava ~2 dias para múltiplas notas
-   **Depois:** Processamento automatizado em ~3-5 segundos por nota
-   **Impacto:** Economia de 95% do tempo + redução de erros humanos

### ✨ Diferenciais

-   ✅ **Instalação Zero Config:** Setup automático de Python e dependências
-   ✅ **Gerenciamento Inteligente de AJAX:** Wait otimizado para carregamento dinâmico
-   ✅ **Sistema de Retry:** Tratamento robusto de falhas temporárias
-   ✅ **Multi-Cliente:** Suporte a diferentes configurações de prestadores
-   ✅ **Performance:** 60% mais rápido que versão inicial
-   ✅ **Logs Detalhados:** Sistema completo de debugging e monitoramento

## 🚀 Instalação

### Método 1: Instalação Automática (Recomendado)

```bash
# Windows
setup_python.bat

# Ou execute manualmente
python instalar_dependencias.py
```

### Método 2: Instalação Manual

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/rpa-notas-fiscais.git
cd rpa-notas-fiscais

# 2. Crie ambiente virtual (opcional mas recomendado)
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows

# 3. Instale dependências
pip install -r requirements.txt
```

## 📊 Estrutura do Projeto

```
rpa-notas-fiscais/
├── 📄 rpa_notas_fiscais.py      # Script principal do RPA
├── ⚙️  setup_python.bat          # Instalador automático (Windows)
├── 🔧 instalar_dependencias.py  # Instalador de dependências Python
├── ▶️  INICIAR_RPA.bat           # Executor rápido (Windows)
├── 📋 requirements.txt          # Dependências do projeto
├── 📊 notas_fiscais.xlsx        # Arquivo de dados (exemplo)
├── 📝 example.py                # Gerador de arquivo exemplo
└── 📖 README.md                 # Este arquivo
```

## 💻 Como Usar

### 1. Prepare seus Dados

Crie um arquivo Excel com as seguintes colunas obrigatórias:

| Coluna         | Descrição            | Exemplo     |
| -------------- | -------------------- | ----------- |
| `Nome_Cliente` | Nome completo        | João Silva  |
| `CPF`          | CPF (só números)     | 12345678901 |
| `Nome_Item`    | Nome do item/serviço | Serviço A   |
| `Valor`        | Valor do serviço     | 150.00      |
| `Data`         | Data (DD/MM/AA)      | 25/09/24    |
| `Cidade`       | Cidade (opcional)    | CIDADE_A    |
| `Endereco`     | Endereço (opcional)  | R Nome, 123 |

**Dica:** Use `python example.py` para gerar um arquivo de exemplo.

### 2. Execute o RPA

#### Modo Windows (Mais Fácil):

```bash
INICIAR_RPA.bat
```

#### Modo Python Direto:

```bash
python rpa_notas_fiscais.py
```

### 3. Siga as Instruções

1. **Selecione o cliente** (Cliente A ou Cliente B)
2. **Faça login** no sistema web quando o navegador abrir
3. **Navegue** até a página de emissão de notas
4. **Pressione ENTER** para iniciar o preenchimento automático

## ⚙️ Configuração

### Mapeamento de Clientes

O sistema suporta múltiplos perfis de cliente com configurações específicas:

```python
# Exemplo de configuração (em rpa_notas_fiscais.py)
mapeamentos_clientes = {
    'cliente_a': {
        'atividade': '508',
        'aliquota': '2,01',
        'municipio_incidencia': 'CIDADE_EXEMPLO',
        # ... outras configurações
    }
}
```

### Personalização

Para adicionar novo cliente, edite o dicionário `mapeamentos_clientes` em `rpa_notas_fiscais.py` e adicione suas configurações específicas.

## 🏗️ Arquitetura Técnica

### Stack Tecnológico

-   **Python 3.8+:** Linguagem principal
-   **Selenium 4.15+:** Automação web
-   **Pandas 2.1+:** Processamento de dados
-   **WebDriver Manager:** Gerenciamento automático do ChromeDriver
-   **openpyxl 3.1+:** Leitura/escrita de Excel

### Principais Componentes

#### 1. Gerenciamento de WebDriver

```python
# Download e configuração automática do ChromeDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
```

#### 2. Wait Inteligente para AJAX

```python
def aguardar_ajax_cpf(self, timeout=8):
    """
    Aguarda carregamento AJAX com múltiplas verificações:
    - Estado de requisições jQuery/PrimeFaces
    - Indicadores de loading
    - Disponibilidade de campos
    """
```

#### 3. Sistema de Retry Robusto

```python
def selecionar_dropdown(self, element_id, value, retry_count=3):
    """
    Múltiplas estratégias de seleção:
    1. Select por value
    2. Select por texto visível
    3. Busca flexível
    4. Click + panel (fallback)
    """
```

## 📈 Performance

### Métricas de Otimização

| Métrica           | Versão 1.0 | Versão 2.0 | Melhoria               |
| ----------------- | ---------- | ---------- | ---------------------- |
| Tempo por nota    | 8-12s      | 3-5s       | 🚀 **60% mais rápido** |
| Taxa de sucesso   | 85%        | 95%        | ✅ **+10% confiável**  |
| Wait AJAX CPF     | 3s fixo    | 0.5-1s     | ⚡ **Inteligente**     |
| Delay entre notas | 2s         | 0.5s       | 🎯 **75% redução**     |

### Logs de Performance

Durante a execução, o sistema exibe métricas em tempo real:

```
⏱️  MONITORAMENTO DE PERFORMANCE:
========================================
📝 Nota 1: 3.2s
📝 Nota 2: 2.8s
📝 Nota 3: 3.5s
...
========================================
📊 RELATÓRIO DE PERFORMANCE
========================================
⏱️  Tempo total: 156.3s
📈 Tempo médio por nota: 3.1s
🚀 Tempo mínimo: 2.8s
⚡ Tempo máximo: 3.5s
```

## 🔧 Resolução de Problemas

### ❌ "Python não encontrado"

```bash
setup_python.bat
```

### ❌ "ChromeDriver não funciona"

```bash
pip install webdriver-manager --upgrade
```

### ❌ "Erro no dropdown"

-   Verifique se o site não mudou a estrutura HTML
-   Ative logs detalhados para debug
-   Execute em modo teste primeiro

### ❌ "CPF não preenche"

-   Aguarde o carregamento completo da página
-   Verifique se o CPF tem exatamente 11 dígitos numéricos
-   Confira se não há campos obrigatórios anteriores não preenchidos

## 📝 Logs e Debug

### Ativando Logs Detalhados

Os logs são salvos automaticamente em `rpa_notas_fiscais.log`:

```python
# Nível de log configurável em setup_logging()
logging.basicConfig(
    level=logging.INFO,  # Altere para DEBUG para mais detalhes
    format='%(asctime)s - %(levelname)s - %(message)s'
)
```

### Interpretando Logs

```
2024-11-02 14:30:15 - INFO - Processando registro 1/10
2024-11-02 14:30:16 - INFO - CPF preenchido: 12345678901
2024-11-02 14:30:17 - INFO - CPF AJAX completo - campos preenchidos automaticamente
2024-11-02 14:30:18 - INFO - Formulário preenchido com sucesso
```

## 🚨 Importante

-   ⚠️ **Execute em modo TESTE primeiro** para validar os dados
-   ⚠️ **Mantenha backup** dos arquivos Excel
-   ⚠️ **Verifique cada nota** antes da emissão definitiva
-   ⚠️ **O site pode mudar** e quebrar a automação (requer manutenção)

## 🛣️ Roadmap

-   [ ] Interface gráfica (GUI) para usuários não-técnicos
-   [ ] Suporte a múltiplos navegadores (Firefox, Edge)
-   [ ] Sistema de notificações (email/Slack)
-   [ ] Dashboard de analytics e relatórios
-   [ ] API REST para integração com outros sistemas
-   [ ] Testes automatizados (pytest)
-   [ ] Docker container para deploy facilitado

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## 👤 Autor

**Bruno Rufatto**

-   GitHub: [@ymist](https://github.com/ymist)
-   LinkedIn: [Bruno Rufatto](https://linkedin.com/in/bruno-rufatto)

⭐ Se este projeto te ajudou, considere dar uma estrela!

**Desenvolvido para resolver problemas reais, não apenas para código bonito.**
