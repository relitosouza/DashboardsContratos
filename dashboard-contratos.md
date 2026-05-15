# Plano de Implementação: Dashboard de Contratos

## Fases da Execução
1. **ANALYSIS**: O usuário deseja um dashboard completo e moderno para analisar uma planilha de contratos (`RESOUZA_Planilha de Gerenciamento de Componentes (7).xlsx`). O dashboard deve agrupar por secretaria, calcular métricas úteis e ter uma interface de alta qualidade.
2. **PLANNING**: Vamos desenvolver uma aplicação web com **Next.js**, **React** e **Tailwind CSS**. A aplicação utilizará a biblioteca **SheetJS (xlsx)** para processar a planilha localmente (no navegador do usuário) e **Recharts** para as visualizações. Isso permite que qualquer nova versão da planilha possa ser importada dinamicamente com um simples "Arrastar e Soltar".
3. **SOLUTIONING**: 
    - `components/Dashboard.tsx`: Componente principal contendo a área de *Drag & Drop* e a lógica de processamento de dados (busca dinâmica de colunas como "Secretaria", "Valor", "Status").
    - `app/layout.tsx` e `app/globals.css`: Configuração do design system focado em *Glassmorphism* (dark mode elegante).
    - `tailwind.config.js`: Definição dos *design tokens* (cores premium, gradientes).
4. **IMPLEMENTATION**: Criação automatizada de toda a infraestrutura base (Next.js scaffold) para que o usuário precise apenas instalar as dependências e iniciar o projeto.
