# Inventário de Máquinas - Dynamo Tecnologia

Sistema desenvolvido para realizar **levantamento técnico de hardware e software** de computadores de forma rápida e padronizada.

O objetivo é facilitar o trabalho de inventário em clientes, gerando automaticamente uma planilha Excel consolidada com todas as máquinas analisadas.

---

# Funcionalidades

✔ Identifica automaticamente:

- Nome da máquina
- Modelo do equipamento
- Fabricante
- Serial da BIOS
- Processador
- Memória RAM
- Tipo de disco (SSD / HD)
- Tamanho do disco
- Antivírus instalado
- Versão do Windows
- Status da licença do Windows

✔ Permite adicionar manualmente:

- Nome do usuário
- Setor
- Sistemas utilizados
- Tipo de conta (Google Drive, OneDrive, etc)
- E-mail da conta
- Chave de licença manual
- Observações

✔ Exportação automática para **Excel**

✔ Planilha **acumulativa**, adicionando cada máquina na próxima linha

✔ Layout organizado para envio ao cliente

---

# Como funciona

1. Execute o programa
2. O sistema coleta automaticamente os dados da máquina
3. Preencha os campos adicionais se necessário
4. Clique em **Adicionar ao Excel**

O sistema irá:

- Criar o arquivo `InventarioMaquinas.xlsx` se não existir
- Ou adicionar a máquina na próxima linha da planilha existente

# Tecnologias utilizadas

- C#
- .NET 8
- Windows Forms
- ClosedXML (exportação Excel)

# Uso em campo

Para realizar inventário em várias máquinas:

Leve o executável em um pendrive

Utilize em cada computador

O Excel será atualizado automaticamente com cada máquina