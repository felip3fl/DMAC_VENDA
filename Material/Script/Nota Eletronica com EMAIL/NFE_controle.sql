

CREATE TABLE [dbo].[NFE_controle](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[danfe_IMPRESSORA] [varchar](100) NULL,
	[danfe_RETORNARESP] [char](1) NULL,
	[email_DESTINATARIO] [varchar](60) NULL,
	[email_ASSUNTO] [varchar](60) NULL,
	[email_MENSAGEM] [varchar](120) NULL,
	[email_EMAILEMITENTE] [varchar](60) NULL,
	[email_NOMEEMITENTE] [varchar](60) NULL,
	[email_ANEXOPDF] [char](3) NULL,
	[email_ANEXOXML] [char](3) NULL,
	[email_ANEXOPROTOCOLO] [char](3) NULL,
	[email_anexoadicional] [char](3) NULL,
	[email_COMPACTADO] [char](3) NULL,
	[email_RETORNARESP] [char](1) NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


