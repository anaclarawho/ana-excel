
📊 PLANILHA DE APOIO À DECLARAÇÃO DE IMPOSTO DE RENDA

Este repositório contém a planilha desenvolvida como projeto prático do curso,
com foco em facilitar a organização das informações necessárias para a
declaração de Imposto de Renda.

A proposta foi construir uma ferramenta funcional, organizada e automatizada,
capaz de reduzir erros e otimizar o processo de preparação da declaração.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 OBJETIVO

Criar uma planilha estruturada para:

✔ Organizar rendimentos  
✔ Controlar despesas dedutíveis  
✔ Automatizar cálculos  
✔ Centralizar informações  
✔ Reduzir retrabalho na hora da declaração  

A ideia foi transformar a planilha em um sistema de apoio anual,
não apenas um arquivo de uso pontual.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🛠 MEU PROCESSO DE DESENVOLVIMENTO

A planilha foi construída em conjunto com o professor durante o curso,
seguindo um processo estruturado e orientado à prática.

Primeiro, defini a estrutura das abas, organizando as informações por categorias.
Depois, modelei as tabelas com campos objetivos e fórmulas automáticas,
garantindo consistência nos cálculos.

Em seguida, apliquei validação de dados para reduzir erros de digitação
e inconsistências.

Na etapa final, foquei na padronização visual,
mantendo uniformidade entre ícones, imagens e elementos gráficos
em todas as abas.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎨 PADRONIZAÇÃO VISUAL COM VBA

Durante o desenvolvimento, percebi que alinhar manualmente os ícones
em cada planilha gerava retrabalho e imprecisão.

O método tradicional era testar posição por posição.
Optei por uma abordagem mais técnica.

Posicionei o ícone na primeira planilha exatamente onde eu queria.
Depois, capturei as coordenadas X e Y utilizando VBA:

? Selection.ShapeRange(1).Left
? Selection.ShapeRange(1).Top

Com os valores anotados, apliquei as mesmas coordenadas
nas demais planilhas com o código abaixo:

------------------------------------------------------------
Sub MoverIconeParaPosicao()
    Dim shp As Shape
    Dim ws As Worksheet
    Dim nomeIconeProcurado As String
    Dim novaPosicaoX As Double
    Dim novaPosicaoY As Double
    
    Set ws = ActiveSheet
    
    nomeIconeProcurado = "Ícone 1"
    
    novaPosicaoX = 100
    novaPosicaoY = 50
    
    For Each shp In ws.Shapes
        If shp.Name = nomeIconeProcurado Then
            shp.Left = novaPosicaoX
            shp.Top = novaPosicaoY
            Exit Sub
        End If
    Next shp
End Sub
------------------------------------------------------------

Essa decisão trouxe:

✔ Precisão no alinhamento  
✔ Padronização visual entre abas  
✔ Economia de tempo  
✔ Redução de retrabalho  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💻 TECNOLOGIAS UTILIZADAS

• Microsoft Excel  
• Fórmulas e funções financeiras  
• Validação de dados  
• VBA para automação visual  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚀 RESULTADO

Uma planilha organizada, automatizada e visualmente consistente,
construída com base no conteúdo aprendido no curso
e aprimorada com soluções técnicas aplicadas durante o desenvolvimento.
