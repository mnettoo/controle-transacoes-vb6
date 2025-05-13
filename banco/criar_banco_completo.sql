-- Criação do banco
CREATE DATABASE TransacoesDB;
GO

USE TransacoesDB;
GO

-- Tabela de transações
CREATE TABLE Transacoes (
    Id_Transacao INT IDENTITY(1,1) PRIMARY KEY,
    Numero_Cartao VARCHAR(16) NOT NULL,
    Valor_Transacao DECIMAL(18,2) NOT NULL CHECK (Valor_Transacao > 0),
    Data_Transacao DATETIME NOT NULL,
    Descricao VARCHAR(255) NOT NULL,
    Status_Transacao VARCHAR(20) NOT NULL CHECK (Status_Transacao IN ('Aprovada', 'Pendente', 'Cancelada'))
);
GO

-- Dados de exemplo
INSERT INTO Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao) VALUES
('1234567812345678', 450.00, '2024-05-10', 'Assinatura mensal', 'Pendente'),
('9876543298765432', 2200.00, '2024-05-11', 'Compra notebook', 'Aprovada'),
('4567891245678912', 1500.50, '2024-04-25', 'Curso online', 'Cancelada'),
('7894561278945612', 80.00, '2024-05-01', 'Uber', 'Pendente'),
('3214567832145678', 999.99, '2024-05-12', 'Compra roupas', 'Aprovada');
GO

-- Scalar Function: fn_CategorizaValor
CREATE FUNCTION fn_CategorizaValor (@valor DECIMAL(18,2))
RETURNS VARCHAR(20)
AS
BEGIN
    DECLARE @categoria VARCHAR(20)

    IF @valor > 2000 SET @categoria = 'Premium'
    ELSE IF @valor >= 1000 SET @categoria = 'Alta'
    ELSE IF @valor >= 500 SET @categoria = 'Média'
    ELSE SET @categoria = 'Baixa'

    RETURN @categoria
END;
GO

-- TVF: fn_TransacoesCategorizadas
CREATE FUNCTION fn_TransacoesCategorizadas (@Data_Inicial DATETIME, @Data_Final DATETIME)
RETURNS TABLE
AS
RETURN
(
    SELECT 
        t.Id_Transacao,
        t.Numero_Cartao,
        t.Valor_Transacao,
        t.Data_Transacao,
        t.Descricao,
        t.Status_Transacao,
        dbo.fn_CategorizaValor(t.Valor_Transacao) AS Categoria
    FROM 
        Transacoes t
    WHERE 
        t.Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
);
GO

-- Stored Procedure: sp_ResumoTransacoes
CREATE PROCEDURE sp_ResumoTransacoes
    @Data_Inicial DATETIME,
    @Data_Final DATETIME,
    @Status_Transacao VARCHAR(20)
AS
BEGIN
    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Quantidade_Transacoes,
        Status_Transacao
    FROM 
        Transacoes
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
        AND Status_Transacao = @Status_Transacao
    GROUP BY 
        Numero_Cartao, Status_Transacao
END;
GO

-- View: vw_ResumoFinanceiro
CREATE VIEW vw_ResumoFinanceiro AS
SELECT 
    Numero_Cartao,
    COUNT(*) AS Quantidade,
    SUM(Valor_Transacao) AS Total,
    AVG(Valor_Transacao) AS Media,
    MIN(Valor_Transacao) AS Menor_Valor,
    MAX(Valor_Transacao) AS Maior_Valor
FROM 
    Transacoes
GROUP BY 
    Numero_Cartao;
GO
