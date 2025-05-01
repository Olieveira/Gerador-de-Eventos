<h1 align="center">Gerador de Eventos</h1>

<p align="center">
  <img src="https://user-images.githubusercontent.com/107584427/207604012-331361e1-2d1a-4ce3-bc89-8fefc53868e8.png" alt="Capa do Projeto" />
</p>

<p align="center">
  Gerador de cronogramas desenvolvido em <strong>VBA (Visual Basic for Applications)</strong>, utilizando as funcionalidades do Excel para criar uma interface dinâmica que permite a geração, exportação e impressão de eventos de forma automatizada e intuitiva.
</p>

---

## 🚀 Funcionalidades

- **Geração dinâmica**  
  Criação de itens do cronograma em tempo de execução via VBA, adaptando-se à quantidade de dados inseridos.

- **Paginação automática**  
  Permite a exibição controlada de itens (8 por página), com navegação e separação entre páginas.

<p align="center">
  <img src="https://user-images.githubusercontent.com/107584427/207616830-ca60db58-6515-4083-b58f-6289cb1589c6.gif" height="350px" />
</p>

- **Exportações disponíveis:**
  - `HTML:` Geração de tabela com largura ajustada ao conteúdo.
  
    <p align="center">
      <img src="https://user-images.githubusercontent.com/107584427/207620333-0166fc04-136e-4ac8-9989-e8b869ba5114.png" height="350px" />
    </p>

  - `PDF:` Layout com espaçamento aprimorado e quebras de linha automáticas.

    <p align="center">
      <img src="https://user-images.githubusercontent.com/107584427/207623054-e40501da-7da0-4588-a346-19f24a9b4604.png" height="350px" />
    </p>

  - `TXT:` Exportação de texto simples, sem cabeçalhos, baseado na estrutura da planilha.

    <p align="center">
      <img src="https://user-images.githubusercontent.com/107584427/207654545-a2adb9ae-e1c9-4d26-a97e-1302b499614b.png" height="350px" />
    </p>

  - `DOC:` Documento Word com margens ajustadas, colagem formatada e cabeçalhos estilizados.

    <p align="center">
      <img src="https://user-images.githubusercontent.com/107584427/207655994-46d11525-77b9-4085-a282-468569ccf987.png" height="400px" />
    </p>

  - `XLSX:` Planilha com área de impressão e alinhamento formatados.

    <p align="center">
      <img src="https://user-images.githubusercontent.com/107584427/207657922-b5319659-0f6f-4561-a255-1f5fc0d97b4a.png" height="400px" />
    </p>

- **Impressão direta**  
  Cria automaticamente uma versão da planilha adaptada para impressão.

- **Importação inteligente**  
  Permite importar arquivos `.xlsx` exportados previamente, mantendo compatibilidade com o modelo.

<p align="center">
  <img src="https://user-images.githubusercontent.com/107584427/207663704-064fef4a-98e2-45f2-a055-1d92629d6b53.gif" height="350px" />
</p>

---

## 📁 Estrutura de Arquivos Exportados

Os arquivos gerados são automaticamente salvos na mesma pasta onde a planilha está localizada, dentro da estrutura:

/Programas/{extensão}

Exemplos:
- `Programas/pdf/cronograma.pdf`
- `Programas/html/cronograma.html`
- `Programas/xlsx/cronograma.xlsx`

A criação das pastas é feita automaticamente caso não existam.

---

## 📄 Exemplo de Cronograma Exportado

| N°  | DESCRIÇÃO             | PARTICIPANTES        | INSTRUMENTOS      | MÍDIA              | PROJEÇÃO     |
|-----|------------------------|-----------------------|--------------------|---------------------|---------------|
| 1 - | Abertura do Evento     | Equipe Organizadora   | Microfone, Caixa   | Link Apresentação   | NENHUM        |
| 2 - | Apresentação Cultural  | Grupo Musical         | Violão, Teclado    | YouTube Link        | SOM           |
| 3 - | Palestra Principal     | Convidado Especial    | Projetor           | Slides              | SOM E VÍDEO   |
| 4 - | Encerramento           | Todos                 | Microfone          | Link Final          | NENHUM        |

---

## 🧩 Estrutura do Código

O projeto foi modularizado para melhor organização e manutenção:

| Componente       | Descrição                                                              |
|------------------|------------------------------------------------------------------------|
| `cadastro`       | Formulário principal com interface para o usuário                      |
| `clsExporter`    | Responsável pelas exportações em diferentes formatos                   |
| `clsFormatter`   | Cuida da estilização das planilhas geradas (títulos, alinhamento etc.) |
| `clsSystem`      | Interações com o sistema: criação de pastas, impressão, abertura de arquivos |

---

## 🛠 Tecnologias Utilizadas

- Microsoft Excel (com suporte a macros - `.xlsm`)
- VBA (Visual Basic for Applications)
- Microsoft Word

---

## 🎯 Objetivo

O projeto visa facilitar a criação padronizada de cronogramas de eventos para contextos profissionais, acadêmicos ou pessoais, promovendo:

- Automação de tarefas manuais no Excel
- Uso de boas práticas em VBA
- Modularização e reaproveitamento de código
- Interação entre planilhas, diretórios e documentos externos
