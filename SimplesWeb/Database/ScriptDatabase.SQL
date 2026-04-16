Create Database SimplesWeb;
Use SimplesWeb;

CREATE TABLE pessoa (
  Id int NOT NULL AUTO_INCREMENT,
  RazaoSocial varchar(80) DEFAULT NULL,
  NomeFantasia varchar(80) DEFAULT NULL,
  Cnpj varchar(18) DEFAULT NULL,
  InscrEstadual varchar(20) DEFAULT NULL,
  InscrMunicipal varchar(20) DEFAULT NULL,
  DataCadastro varchar(20) DEFAULT NULL,
  Email varchar(80) DEFAULT NULL,
  Celular varchar(15) DEFAULT NULL,
  Usuario varchar(20) DEFAULT NULL,
  Senha varchar(12) DEFAULT NULL,
  IeDestinatario varchar(20) DEFAULT NULL,
  PRIMARY KEY (Id)
);

INSERT INTO pessoa VALUES (1,'Pessoa 1','Pessoa 1','00.000.000/0000-00','','','','email@email.com.br','(11) 00000-0000','','','Isento');
INSERT INTO pessoa VALUES (2,'Pessoa 2 ','Pessoa 2 ','00.000.000/0000-00','','','','email2@gmail.com.br','(00) 00001-1000','','','Isento');
INSERT INTO pessoa VALUES (3,'Pessoa 3','Pessoa 3','00.000.000/0000-00','','','','email3@email.com.br','(00) 10000-0001','','','Isento');
INSERT INTO pessoa VALUES (4,'Pessoa 4','Pessoa 4','11.111.111/1111-11','','','','email4@email.com.br','(01) 00001-0001','','','Isento');
INSERT INTO pessoa VALUES (5,'Pessoa 5','Pessoa 5','00.000.000/0000-00','','','','email5@email.com.br','(11) 01100-0110','SeuUsuario','SuaSenha','Isento');
INSERT INTO pessoa VALUES (6,'Pessoa 6','Pessoa 6','00.000.000/0000-00','','','','email6@email.com.br','(00) 00000-0000','','','Isento');
INSERT INTO pessoa VALUES (7,'Pessoa 7','Pessoa 7','00.000.000/0000-00','','','','email7@email.com.br','(00) 00000-0000','','','Isento');
INSERT INTO pessoa VALUES (8,'Pessoa 8','Pessoa 8','00.000.000/0000-00','','','','email@email.com.br','(00) 00000-1111','','','Isento');
INSERT INTO pessoa VALUES (9,'Pessoa 9 ','Pessoa 9 ','00.000.000/0000-00','','','','email9@email.com.br','(00) 11111-0000','','','Isento');
INSERT INTO pessoa VALUES (10,'Pessoa 10 ','Pessoa 10 ','00.000.000/0000-00','','','','email10@email.com.br','(00) 11111-0000','','','Isento');
