
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

Durante o desenvolvimento, o professor ensinou a ajustar os ícones inserindo valores de X e Y diretamente no código e ir testando até encontrar a posição correta. Ou seja, alterar os números, executar, verificar, corrigir e repetir até alinhar.

Percebi que esse processo exigia várias tentativas e tomava mais tempo.

Então adotei uma abordagem diferente.

Primeiro, posicionei o ícone manualmente na primeira planilha exatamente onde eu queria que ele ficasse.

Depois, utilizei o VBA para capturar automaticamente as coordenadas exatas daquela posição:

? Selection.ShapeRange(1).Left
? Selection.ShapeRange(1).Top

Esses comandos retornaram os valores reais de X e Y do ícone já alinhado.

Com essas coordenadas anotadas, bastou copiar os valores para o código do professor e aplicar nas demais abas, garantindo que todos os ícones ficassem perfeitamente alinhados sem precisar testar posição por posição.

Essa solução trouxe:

✔ Precisão no alinhamento
✔ Economia de tempo
✔ Padronização entre abas
✔ Processo mais inteligente de ajuste visual

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
