# Instruções de uso

O projeto tem o intuito de avaliar meu conhecimentos na linguagem VB6, com a criação de um sistema simples do zero.

## Funcionalidades:

O sistema inclui as seguintes telas:
- Cadastro de corretores
- Cadastro de clientes
- Consulta de clientes


## Pré-requisitos
Antes de iniciar, verifique se você possui:
- Visual Basic 6 (VB6) instalado na máquina
- SQL Server instalado na máquina

## Instruções

### 1. Clonagem do repositório

Para clonar este repositório, execute o comando abaixo em seu terminal:

```
git clone https://github.com/LuscaMD/CorretagemFacil.git
```

### 2. Ambientação do Banco de dados

#### 2.1 Criação do base de dados no seu servidor
Para criar uma base de dados no SQL Server, siga estes passos:
- 1° Abrir o SQL Server >> Menu Superior >> View >> Object Explorer
- 2° Depois de conectado no servidor >> Botão direito em "Databases" >> New Database
- 3° Em "Database name" colocar o nome da base que você quer criar

#### 2.2 Executar scripts SQL
Para configurar o banco de dados com as tabelas necessárias e dados iniciais, execute o script SQL localizado em "CorretagemFacil/SQL/Scripts.sql" no SQL Server.

#### 2.3 Trocar a Connection String do aplicativo
Para trocar a connection string do aplicativo devemos seguir os seguintes passos:
- 1° Ir em "CorretagemFacil/Aplicativo/BancoDeDados.bas" 
- 2° No método "PreencheConnetionString" trocar os valores de "Data Source" e "Initial Catalog" para o seu servidor e sua base de dados, respectivamente.

```vb6
' Padrão
Public Sub PreencheConnetionString()
  pub_str_ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-KULQ16T\SQLEXPRESS;Initial Catalog=ViceriSeidor;Integrated Security=SSPI;"
End Sub

' Sua versão
Public Sub PreencheConnetionString()
  pub_str_ConnectionString = "Provider=SQLOLEDB;Data Source=NomeServidor;Initial Catalog=NomeBase;Integrated Security=SSPI;"
End Sub
```