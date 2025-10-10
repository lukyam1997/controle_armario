# Controle de Armários

Aplicação em Google Apps Script para organizar o controle dos armários de pertences em unidades hospitalares do SUS.

## Funcionalidades

- Registro rápido de novos armários ocupados com indicação de unidade e perfil (visitante, acompanhante, equipe).
- Monitor visual de todos os armários cadastrados com indicação de disponibilidade, itens guardados e responsável.
- Filtros globais por unidade e, após a seleção da unidade, por perfil de usuário (visitante ou acompanhante) para refinar o monitor e a tabela.
- Painel em tempo real com contagem de registros ativos, finalizados, ocupação percentual e distribuição por perfil.
- Sincronização automática dos armários cadastrados na planilha para manter o monitor sempre completo.
- Ações de encerramento, reabertura e exclusão diretamente da tabela.
- Sugestão automática das unidades já cadastradas para padronizar o preenchimento.

## Planilha

Ao publicar o Web App, duas abas principais serão criadas automaticamente (caso não existam):

### Aba `ControleArmarios`

Usada para armazenar o histórico de movimentações. Colunas:

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

### Aba `ArmariosMonitor`

Responsável por manter o catálogo dos armários exibidos no monitor. Colunas:

1. Unidade
2. Armario
3. Descricao
4. Capacidade
5. Observacoes

A aba é preenchida automaticamente com exemplos iniciais e recebe novos armários sempre que um registro é criado para uma combinação de unidade e armário inédita. É possível complementar manualmente as descrições, capacidades e observações para enriquecer o monitor.

A partir dessas abas é possível acompanhar e auditar todos os lançamentos.

## Uso

1. Abra o projeto no Google Apps Script vinculado a uma planilha Google.
2. Publique o Web App concedendo acesso às equipes responsáveis.
3. Utilize o formulário do painel para novos registros e monitore as movimentações em tempo real.

