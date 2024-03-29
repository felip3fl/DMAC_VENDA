SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[faixapremioge_novo](
	[FPG_Plano] [numeric](18, 0) NOT NULL,
	[FPG_CodigoFaixa] [numeric](18, 0) NOT NULL,
	[FPG_FaixaInicial] [float] NULL,
	[FPG_FaixaFinal] [float] NULL,
	[FPG_PremioLiquido] [float] NULL,
	[FPG_Remuneracao] [float] NULL,
	[FPG_PISCOFINS] [float] NULL,
	[FPG_IOF] [float] NULL,
	[FPG_Premio] [float] NULL,
 CONSTRAINT [PK_faixapremioge_novo] PRIMARY KEY CLUSTERED 
(
	[FPG_Plano] ASC,
	[FPG_CodigoFaixa] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(1 AS Numeric(18, 0)), 0, 50, 1.75, 2.59, 0.21, 0.34, 4.89)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(2 AS Numeric(18, 0)), 50.01, 100, 4.46, 6.64, 0.54, 0.86, 12.5)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(3 AS Numeric(18, 0)), 100.01, 150, 6.52, 9.69, 0.79, 1.25, 18.26)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(4 AS Numeric(18, 0)), 150.01, 200, 9.22, 13.73, 1.12, 1.78, 25.84)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(5 AS Numeric(18, 0)), 200.01, 250, 10.75, 16, 1.3, 2.07, 30.12)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(6 AS Numeric(18, 0)), 250.01, 300, 12.78, 19.02, 1.55, 2.46, 35.81)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(7 AS Numeric(18, 0)), 300.01, 400, 18.15, 27.01, 2.2, 3.5, 50.86)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(8 AS Numeric(18, 0)), 400.01, 500, 23.65, 35.18, 2.87, 4.55, 66.25)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(9 AS Numeric(18, 0)), 500.01, 700, 29.03, 43.19, 3.52, 5.59, 81.33)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(24 AS Numeric(18, 0)), CAST(10 AS Numeric(18, 0)), 700.01, 9999, 50.37, 74.93, 6.11, 9.7, 141.11)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(1 AS Numeric(18, 0)), 0, 50, 2.81, 4.18, 0.34, 0.54, 7.87)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(2 AS Numeric(18, 0)), 50.01, 100, 7.16, 10.64, 0.87, 1.38, 20.05)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(3 AS Numeric(18, 0)), 100.01, 150, 10.67, 15.88, 1.29, 2.05, 29.9)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(4 AS Numeric(18, 0)), 150.01, 200, 15.57, 23.17, 1.8900000000000001, 3, 43.63)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(5 AS Numeric(18, 0)), 200.01, 250, 17.72, 26.36, 2.15, 3.41, 49.64)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(6 AS Numeric(18, 0)), 250.01, 300, 21.37, 31.8, 2.59, 4.12, 59.88)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(7 AS Numeric(18, 0)), 300.01, 400, 28.55, 42.45, 3.46, 5.5, 79.96)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(8 AS Numeric(18, 0)), 400.01, 500, 38.04, 56.58, 4.61, 7.32, 106.56)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(9 AS Numeric(18, 0)), 500.01, 700, 47.81, 71.12, 5.8, 9.21, 133.94)
INSERT [dbo].[faixapremioge_novo] ([FPG_Plano], [FPG_CodigoFaixa], [FPG_FaixaInicial], [FPG_FaixaFinal], [FPG_PremioLiquido], [FPG_Remuneracao], [FPG_PISCOFINS], [FPG_IOF], [FPG_Premio]) VALUES (CAST(36 AS Numeric(18, 0)), CAST(10 AS Numeric(18, 0)), 700.01, 9999, 81.9, 121.82, 9.93, 15.77, 229.42)
