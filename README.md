# Excel Automation with Python

Automação desenvolvida em Python para atualização automática de uma planilha Excel utilizada no processo de produção.

O script abre uma planilha `.xlsm`, atualiza todas as conexões de dados, ajusta a formatação das datas e gera automaticamente uma nova versão do arquivo em `.xlsx`.

Além disso, o script foi configurado para execução automática utilizando o **Agendador de Tarefas do Windows**, eliminando a necessidade de atualização manual recorrente.

---

# Problema

A equipe precisava atualizar manualmente uma planilha sempre que surgia uma demanda urgente da produção.

Esse processo exigia:

- interromper outras atividades
- abrir o Excel manualmente
- atualizar os dados
- corrigir formatações
- salvar uma nova versão do arquivo

Esse fluxo consumia tempo e estava sujeito a erros manuais.

---

# Solução

Foi desenvolvido um script em **Python** utilizando a biblioteca **win32com** para automatizar todo o processo.

O script realiza automaticamente:

- abertura da planilha Excel
- atualização das conexões de dados
- cálculo de queries assíncronas
- correção da formatação de datas
- geração de um novo arquivo `.xlsx`
- fechamento seguro da aplicação Excel

A automação também foi integrada ao **Windows Task Scheduler**, permitindo execução automática semanal.

---

# Tecnologias utilizadas

- Python
- win32com
- Microsoft Excel
- Windows Task Scheduler

---

# Como funciona

Fluxo da automação:

1. Abre o arquivo `.xlsm`
2. Atualiza todas as conexões de dados
3. Aguarda conclusão das consultas
4. Ajusta formatação de colunas de data
5. Salva o arquivo atualizado
6. Gera uma cópia `.xlsx`
7. Fecha o Excel automaticamente


