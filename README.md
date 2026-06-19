# 🏥 Smart Scheduling Hub - HCPA (Sistema de Agendamento Cirúrgico)

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![HTML5/CSS3](https://img.shields.io/badge/Frontend_Custom-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![Gestão de TI](https://img.shields.io/badge/IT_Management-Orchestration-10a37f?style=for-the-badge)

Um sistema completo de gestão hospitalar construído 100% dentro do ecossistema Google Workspace. Este projeto transformou uma rotina baseada em planilhas estáticas em uma **Aplicação Web Single Page (SPA)** dinâmica, com controle de acesso rigoroso, integrações com Google Drive e painéis visuais intuitivos.

---

## 🚨 O Problema: O Caos da Planilha Estática
Antes deste sistema, o agendamento de cirurgias de convênio do Hospital de Clínicas de Porto Alegre (HCPA) dependia de uma planilha crua e extensa. 
* **Baixa Usabilidade (UX):** Dificuldade para visualizar horários vagos e alocação de equipes.
* **Risco de Segurança:** Qualquer usuário com acesso à planilha poderia, acidentalmente, apagar ou alterar dados críticos.
* **Falta de Integração:** Anexos (prontuários) e comunicações (WhatsApp) eram feitos em plataformas totalmente separadas, gerando retrabalho.

<img width="1241" height="778" alt="3" src="https://github.com/user-attachments/assets/9966c830-c1b0-4632-97a7-1294d62f56e0" />
---

## 💡 A Solução: Arquitetura de Aplicação no Google Apps Script
Utilizando Engenharia de Prompt e uma visão focada em **Gestão de TI e Arquitetura de Soluções**, orquestrei um sistema em GAS (Google Apps Script) que atua como um Backend/API, servindo uma interface front-end moderna (HTML/CSS/JS) diretamente sobre o Google Sheets.

### ✨ Principais Funcionalidades (Features)

#### 1. Front-end Dinâmico (Visão do Dia)
Transformação dos dados da planilha em **Cards interativos** e um gráfico de **Timeline (Swimlanes)** por salas cirúrgicas. 
* Visualização instantânea de conflitos de horário.
* Filtros rápidos de status (Autorizado, Pendente, Negado).

<img width="1241" height="778" alt="1" src="https://github.com/user-attachments/assets/ffb0e43b-d11d-42cc-8626-c9313ba0aecf" />

#### 2. Controle de Acesso Baseado em Funções (RBAC)
Criação de um sistema de permissões dinâmico que "blinda" o front-end e o back-end:
* **Super Admin / Agendamento:** Acesso total (CRUD).
* **Autorização:** Acesso a campos financeiros e técnicos.
* **Guichê/Recepção:** Apenas alteração de status de chegada (Check 1 e 2).
* **Leitores:** Apenas visualização.

#### 3. Integração Direta com Google Drive e WhatsApp
* **Drive API:** Upload de anexos direto do modal do paciente. O sistema cria automaticamente pastas padronizadas (`Matrícula - Paciente / Data - Procedimento`) e salva o link na base.
* **WhatsApp Gateway:** Geração de mensagens automatizadas (Checklist) baseadas em templates inteligentes, calculando automaticamente as horas de chegada conforme a cirurgia, com link direto para o WhatsApp Web.

<img width="1241" height="778" alt="2" src="https://github.com/user-attachments/assets/371d1231-3d16-45e0-a671-669655e9ef55" />

---

## ⚙️ Destaques Técnicos (Under the Hood)
Como *Vibecoder*, o foco foi garantir que o sistema não apenas funcionasse, mas fosse escalável para múltiplos usuários simultâneos no hospital:

* **Gerenciamento de Concorrência (`LockService`):** Implementação de travas de segurança. Se duas pessoas tentarem editar o mesmo agendamento, o sistema gerencia a fila para evitar corrupção de dados.
* **Otimização de Performance (`CacheService`):** Os dados do dia são cacheados na nuvem do Google por 20 segundos. Isso evita que picos de acessos simultâneos excedam as cotas de leitura da API do Sheets, mantendo a interface rápida.
* **Modo "Detetive" de Deleção:** Scripts de fallback que identificam células protegidas antes de permitir o cancelamento/movimentação de um paciente, evitando falhas silenciosas.

---

## 🚀 Impacto Gerado
* **Redução drástica do tempo de agendamento e reagendamento** através de formulários modais integrados.
* **Zero exclusões acidentais** de dados críticos, graças ao bloqueio total da planilha nativa (usuários operam apenas via interface gráfica).
* **Centralização da Informação:** Anexos, status de recepção, informações clínicas e financeiras, tudo a 1 clique de distância.
