# 📝 Automatizador de Planilhas Excel - KITDOC

Este projeto é um **automatizador de planilhas Excel** desenvolvido em **Python** usando a biblioteca **pandas**.  
O projeto foi pensado para a empresa em que trabalhei, com o objetivo de **reduzir o tempo de submissão das informações necessárias para a criação de contratos no sistema**, de **20/30 dias** para apenas **algumas horas**.

Atualmente, o script preenche automaticamente a **coluna de rubrica** da planilha, identificando a sigla presente na última palavra da coluna `descricao`.  
No futuro, ele será expandido para preencher também **produto, kit, identidade, status e setor**.

## ⚡ Funcionalidades atuais:

- 🟢 Leitura da planilha Excel (`MODELO_LOTE_EXEMPLO.xlsx`);
- 🟢 Identificação da **última palavra da coluna `descricao`** como sigla de rubrica;
- 🟢 Preenchimento da coluna `rubrica` com o valor correspondente do dicionário `rubricas`;
- 🟢 Marcação de células inválidas ou não reconhecidas como `#ERRO#`;
- 🟢 Salvamento da planilha preenchida em um novo arquivo (`MODELO_LOTE - PREENCHIDO.xlsx`);

## 📊 Campos da planilha:

| Campo          | Descrição                                     |
|----------------|-----------------------------------------------|
| 📝 descricao      | Nome do item / descrição (onde está a sigla) |
| 💳 rubrica        | Valor da rubrica definido pelo dicionário   |
| 📦 produto        | Preenchimento automático futuro             |
| 🎁 kit            | Preenchimento automático futuro             |
| 🆔 identidade     | Preenchimento automático futuro             |
| ✅❌ status        | Preenchimento automático futuro             |
| 🏢 setor          | Preenchimento automático futuro             |

> Por enquanto, apenas `rubrica` é preenchido automaticamente. Os demais campos serão automatizados em versões futuras.

## 🚀 Como usar:

1. Instale o **Python 3.13** (ou superior)  
2. Instale a biblioteca pandas:

```bash
pip install pandas
