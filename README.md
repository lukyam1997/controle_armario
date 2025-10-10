# Controle de Armários

Aplicação em Google Apps Script para organizar o controle dos armários de pertences em unidades hospitalares do SUS.

## Funcionalidades

- Registro rápido de novos armários ocupados com indicação de unidade e perfil (visitante, acompanhante, equipe).
- Painel em tempo real com contagem de registros ativos, finalizados e distribuição por perfil.
- Ações de encerramento, reabertura e exclusão diretamente da tabela.
- Sugestão automática das unidades já cadastradas para padronizar o preenchimento.

## Planilha

Ao publicar o Web App, uma aba chamada `ControleArmarios` será criada (caso não exista) com as seguintes colunas:

1. Registrado em
2. Armario
3. Unidade
4. Perfil
5. Responsavel
6. Paciente
7. Contato
8. Itens Guardados
9. Status
10. Encerrado em
11. Observacoes

A partir dessa aba é possível acompanhar e auditar todos os lançamentos.

## Uso

1. Abra o projeto no Google Apps Script vinculado a uma planilha Google.
2. Publique o Web App concedendo acesso às equipes responsáveis.
3. Utilize o formulário do painel para novos registros e monitore as movimentações em tempo real.

