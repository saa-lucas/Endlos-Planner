# ⏳ Endlos Planner | Sistema Híbrido de Mapeamento de Tempo

O **Endlos Planner** é um sistema de monitorização e diagnóstico de rotina projetado para transformar uma Google Sheet (o *"Daten"*) num Dashboard Analítico em Tempo Real. 

Ele atua como um supervisor implacável do seu tempo, quantificando foco, lazer, saúde física/mental e distrações para gerar um relatório executivo de performance diária.

---

## 🚀 Arquitetura Reativa (Google Sheets como State Machine)

O grande desafio de construir aplicações no Google Apps Script é a natureza assíncrona e o excesso de *caching* do Google Sheets. Para garantir consistência no banco de dados e fluidez na interface (Sidebar e Dashboard HTML), o sistema foi desenhado sob os princípios de **Single Source of Truth (SSOT)** e **Concorrência Protegida**.

### 1. The Mutex Engine (`LockService`)
Problema comum: Copiar e colar dezenas de blocos de horas de uma vez ativaria dezenas de gatilhos `onEdit` paralelos, colapsando a planilha sob múltiplas repinturas (`flush`).
* **Solução:** Implementação de um *Thread Lock* (Mutex). Qualquer edição dispara um `Utilities.sleep(1500)` de *debounce*. Apenas a *thread* sobrevivente obtém a chave de execução do `LockService`, forçando a serialização. Nunca ocorrem repinturas simultâneas.

### 2. O System Config Builder (Sidebar Reativa)
A configuração de cores, modos de operação (Entrada/Saída/Desperdício) e agrupamento de ações foi abstraída do código rígido e passada para o controlo do utilizador através da aba `Palette Entry`.
* O Frontend em HTML não guarda estado (Stateless UI).
* Utiliza **Polling Passivo**: a UI interroga a planilha de 5 em 5 segundos, verificando o `SYS_VERSION`. Se detetar uma alteração estrutural feita por outra aba, a UI invalida o seu cache local e reconstrói a árvore de dependências no ecrã automaticamente.

### 3. Fila de Ações Órfãs (Tratamento de Exceções)
Se o utilizador inserir uma ação crua na linha do tempo (aba `Organisieren`) sem definir o seu "Impacto/Modo" no sistema, o *backend* interceta.
* O script atualiza o `SYS_VERSION` e injeta a ação num array pendente.
* A Sidebar bloqueia-se com um Modal Forçado (Overlay escuro) e exige que o utilizador categorize a ação antes de continuar a editar a planilha, blindando a consistência dos relatórios do Dashboard.

---

## 📊 O Dashboard de Inteligência

A UI Analítica converte as horas registadas na base de dados (`Daten`) numa matriz de produtividade, calculando os seguintes eixos:

1. **Performance Score (0-100):** Cálculo da Taxa Pura de Execução vs. Distração.
2. **Base Biológica (0-100):** Cruzamento de horas de Sono (necessidade de ~6.5h mínimas) + Saúde Física (mínimo exigido de 2h/semana) + Saúde Mental.
3. **Smart Insights:** O sistema atua como um consultor rude, entregando diretrizes diretas com base no défice detetado. (Ex: "Sistema invertido: Elimine a distração amanhã", "Risco de burnout iminente devido a sono").

> Adicionalmente, possui uma renderização do estilo *"Spotify Wrapped"*, gerando *cards* em ecrã inteiro com a progressão da sua semana.

---

## 🛠️ Tecnologias Utilizadas
* **Backend:** Google Apps Script (ES6+)
* **Database:** Google Sheets (`Daten` / `Palette Entry` / `Organisieren`)
* **Frontend:** HTML5, CSS3 Nativo, Vanilla JavaScript
* **Visualização de Dados:** Chart.js + chartjs-plugin-datalabels

---

## 💻 Autor
Construído e otimizado por **Lucas Sá** (@saa-lucas).