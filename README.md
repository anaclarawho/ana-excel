# 📊 Planilha de Apoio à Declaração de Imposto de Renda

Este repositório contém a planilha desenvolvida como projeto prático do curso, com o objetivo de facilitar a organização das informações necessárias para a declaração de Imposto de Renda.

A proposta foi criar uma ferramenta simples, funcional e organizada, permitindo reunir os dados ao longo do ano e reduzir erros no momento da declaração.

---

## 🎯 Objetivo do Projeto

- Organizar rendimentos  
- Controlar despesas dedutíveis  
- Estruturar informações por categoria  
- Automatizar cálculos  
- Reduzir erros no preenchimento  
- Economizar tempo na entrega da declaração  

A planilha funciona como um sistema de organização prévia. Os dados são registrados ao longo do ano, evitando acúmulo de documentos no período da declaração.

---

## 🛠 Processo de Desenvolvimento

A planilha foi construída em conjunto com o professor durante o curso, seguindo etapas estruturadas:

### 1️⃣ Estruturação das Abas
- Separação por categorias de informação  
- Organização lógica dos dados  

### 2️⃣ Modelagem das Tabelas
- Criação de campos objetivos  
- Aplicação de fórmulas automáticas  
- Padronização de entradas  

### 3️⃣ Validação e Organização
- Controle de consistência de dados  
- Redução de erros de digitação  

### 4️⃣ Padronização Visual
- Layout limpo  
- Ícones organizados  
- Interface uniforme entre abas  

---

## 🎨 Padronização de Ícones com VBA

Para manter todos os ícones e imagens exatamente alinhados entre as planilhas, foi utilizada automação com VBA.

### Método aplicado

1. Inserir o ícone na primeira planilha  
2. Posicionar manualmente no local desejado  
3. Capturar as coordenadas X e Y  
4. Aplicar automaticamente essas coordenadas nas demais planilhas  

Esse processo substituiu o método manual de testar posições repetidamente, tornando o trabalho mais preciso, mais rápido e mais organizado.

---

## 📍 Como capturar as coordenadas do ícone

Após posicionar o ícone manualmente:

1. Selecione o ícone no Excel  
2. Pressione `Alt + F11` para abrir o Editor VBA  
3. Pressione `Ctrl + G` para abrir a Janela Imediata  
4. Digite os comandos abaixo e pressione Enter  

```vba
? Selection.ShapeRange(1).Left
? Selection.ShapeRange(1).Top

Left retorna a posição X

Top retorna a posição Y

Os valores são exibidos em pontos

Anote esses valores para reutilizar nas demais planilhas.

🔁 Código para aplicar a mesma posição nas outras planilhas
Sub MoverIconeParaPosicao()
    Dim shp As Shape
    Dim ws As Worksheet
    Dim nomeIconeProcurado As String
    Dim novaPosicaoX As Double
    Dim novaPosicaoY As Double
    
    Set ws = ActiveSheet
    
    nomeIconeProcurado = "Ícone 1" ' Nome exato do ícone
    
    novaPosicaoX = 100 ' Substitua pelo valor de Left anotado
    novaPosicaoY = 50  ' Substitua pelo valor de Top anotado
    
    For Each shp In ws.Shapes
        If shp.Name = nomeIconeProcurado Then
            shp.Left = novaPosicaoX
            shp.Top = novaPosicaoY
            Exit Sub
        End If
    Next shp
End Sub
💻 Tecnologias Utilizadas

Microsoft Excel

Fórmulas e funções financeiras

Validação de dados

VBA para automação e padronização visual

🚀 Resultado Final

Uma planilha organizada, automatizada e visualmente uniforme, pronta para apoiar a organização das informações para a declaração de Imposto de Renda.

Este projeto representa aplicação prática do conteúdo aprendido no curso, com foco em eficiência, organização e melhoria de processo.
