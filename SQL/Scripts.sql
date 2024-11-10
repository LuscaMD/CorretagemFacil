CREATE TABLE Cadastros
(
	pk_int_Cadastro INT IDENTITY PRIMARY KEY,
	str_Nome VARCHAR(50) NOT NULL,
	str_CPF VARCHAR(14) NOT NULL,
	str_Endereco VARCHAR(50) NULL,
	fk_int_IdUF INT NULL,
	fk_int_IdCidade INT NULL,
	bit_Ativo BIT NULL DEFAULT 0,
	bit_Corretor BIT NULL DEFAULT 0,
	int_CodCorretor INT NULL
)

CREATE TABLE Estados
(
	pk_int_IdUF INT IDENTITY PRIMARY KEY,
	str_UF VARCHAR(2) NOT NULL,
)

CREATE TABLE Cidades
(
	pk_int_IdCidade INT IDENTITY PRIMARY KEY,
	fk_int_IdUF INT NOT NULL,
	str_NomeCidade VARCHAR(50) NOT NULL
)


-- Insert dos estados
INSERT INTO Estados VALUES('RJ')
INSERT INTO Estados VALUES('SP')

-- Insert das cidades
-- RJ
INSERT INTO Cidades VALUES(1, 'Rio de Janeiro')
INSERT INTO Cidades VALUES(1, 'Duque de Caxias')
-- SP 
INSERT INTO Cidades VALUES(2, 'Ribeirao Preto')
INSERT INTO Cidades VALUES(2, 'Jundiai')
INSERT INTO Cidades VALUES(2, 'Campinas')