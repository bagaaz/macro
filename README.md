# 📘 Macro VBA - Cálculo de Trabalho por Período

Este projeto contém uma macro VBA que realiza cálculos de trabalho baseados em um período definido (Dias, Semanas, Meses, Anos) e distribui os resultados em diferentes períodos: **1 ano (365 dias), 3 anos (1095 dias) e 5 anos (1825 dias)**.

---

## 📌 Como Instalar e Executar

### **1️⃣ Abrir o Editor VBA**
1. **No Excel**, pressione `ALT + F11` para abrir o Editor VBA.

### **2️⃣ Criar um Módulo**
1. No Editor VBA, clique em **Inserir > Módulo**.
2. Um novo módulo será criado no seu projeto VBA.

### **3️⃣ Copiar o Código do Arquivo**
1. **Abra o arquivo `macro.vba`** no seu editor de texto preferido.
2. **Copie todo o conteúdo** do arquivo `macro.vba`.
3. **Cole o código no módulo VBA** que foi criado no Editor VBA.

---

### **4️⃣ Executar a Macro**
1. **Pressione `ALT + F8`** no Excel.
2. Selecione **`PreencherTrabalho`**.
3. Clique em **Executar**.

---

## 📌 O que a macro faz?

### **1️⃣ Processamento dos dados**
- **Lê os valores das colunas:**
    - `E3:E11` → **Ciclo**
    - `F3:F11` → **Unidade**
    - `N3:N11` → **Trabalho**

### **2️⃣ Conversão de períodos**
- Converte os valores do **Ciclo** para **dias** com base na Unidade.
- Regras aplicadas:
    - `"D", "DIA"` → Mantém o valor original.
    - `"S", "SEMANA"` → Multiplica por **7**.
    - `"M", "MES"` → Multiplica por **30**.
    - `"ANO"` → Multiplica por **365**.

### **3️⃣ Cálculo dos períodos**
- Se o valor convertido for válido, aplica os seguintes cálculos:
    - **Para 1 ano (365 dias)** → `(365 / resultado) * Trabalho`
    - **Para 3 anos (1095 dias)** → `(1095 / resultado) * Trabalho`
    - **Para 5 anos (1825 dias)** → `(1825 / resultado) * Trabalho`

- **Os valores são arredondados para números inteiros.**

### **4️⃣ Preenchimento dos resultados**
- **Coluna O** → Resultado para **1 Ano**.
- **Coluna P** → Resultado para **3 Anos**.
- **Coluna Q** → Resultado para **5 Anos**.
- Se houver erro, insere **`-1`** na célula correspondente.

---

## 📌 Colunas afetadas pela macro

| Coluna | Dados manipulados |
|--------|------------------|
| E3:E11 | Período (Ciclo) |
| F3:F11 | Unidade de Tempo |
| N3:N11 | Valor de Trabalho |
| O3:O11 | Cálculo para **1 Ano (365 dias)** |
| P3:P11 | Cálculo para **3 Anos (1095 dias)** |
| Q3:Q11 | Cálculo para **5 Anos (1825 dias)** |

---

## 📌 Conclusão
Essa macro VBA automatiza o cálculo de trabalho considerando diferentes períodos e garante que os valores sejam corretamente preenchidos na planilha. 🚀

Agora, basta seguir os passos e executar no seu Excel! 😊
