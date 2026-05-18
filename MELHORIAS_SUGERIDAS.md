# Melhorias sugeridas para o PELOG WEB

Data do registro: 14/05/2026

## Lista de melhorias

1. Separar o `app.py` em modulos menores: rotas, banco, autenticacao, programacao, relatorios e utilitarios.
2. Remover usuarios e senhas fixos no codigo, levando isso para uma tabela `usuarios` no banco.
3. Criptografar senhas com hash seguro, por exemplo usando `werkzeug.security`.
4. Trocar a `SECRET_KEY` padrao e exigir que ela venha de variavel de ambiente.
5. Adicionar protecao CSRF nos formularios.
6. Criar uma camada de banco separada para evitar SQL misturado diretamente nas rotas.
7. Adicionar migrations de banco com Flask-Migrate/Alembic.
8. Melhorar tratamento de erro no banco com `try/except`, rollback e mensagens amigaveis.
9. Evitar SQL dinamico com nome de tabela sem uma lista permitida.
10. Padronizar permissoes por decoradores como `@login_required`, `@admin_required` e `@encarregado_required`.
11. Validar melhor entradas de formularios: placa, CPF, datas, numeros, notas fiscais e arquivos Excel.
12. Normalizar dados de placa e CPF antes de salvar.
13. Criar paginacao nas tabelas grandes.
14. Melhorar filtros de busca por status, setor, doca, usuario, transportadora e periodo.
15. Evitar repeticao entre CD e CROSS reaproveitando funcoes e componentes.
16. Criar templates base, como `base.html`, para cabecalho, logo, Bootstrap, botoes e estilos comuns.
17. Mover CSS e JavaScript para arquivos separados em `static/css` e `static/js`.
18. Melhorar responsividade mobile, principalmente nas telas de portaria e encarregado.
19. Adicionar confirmacoes visuais mais claras para autorizar, iniciar doca, finalizar e limpar programacao.
20. Proteger melhor acoes destrutivas, como excluir e limpar programacao.
21. Criar log de auditoria para importacao, edicao, exclusao, autorizacao e finalizacao.
22. Adicionar testes automatizados para login, permissoes, fluxo do caminhao, importacao e exportacao.
23. Adicionar arquivo `.env.example` documentando `DATABASE_URL`, `SECRET_KEY` e outras variaveis.
24. Adicionar README tecnico com instalacao, banco, usuarios, execucao local e deploy.
25. Melhorar importacao de Excel com validacao de colunas obrigatorias e erros por linha.
26. Evitar duplicidade na importacao de programacoes.
27. Melhorar exportacoes Excel com filtros aplicados, data de geracao, abas e cabecalhos padronizados.
28. Criar dashboard operacional com totais por status, setor e tempo medio.
29. Melhorar tela de TV com destaque para atrasos, tempo em doca, prioridade e separacao CD/CROSS.
30. Adicionar controle de sessao mais robusto, com expiracao e logout automatico.
31. Configurar ambiente de producao com mais seguranca, sem debug e com variaveis obrigatorias.
32. Organizar estrutura de pastas: `routes/`, `services/`, `repositories/`, `templates/`, `static/`, `config.py`, `extensions.py`.
33. Adicionar mensagens flash padronizadas.
34. Criar constantes para status como `aguardando`, `autorizado`, `na_doca` e `finalizado`.
35. Criar API JSON mais consistente para `/api/portaria`, `/dados_tv` e `/api/tv`.

## Prioridade inicial recomendada

1. Separar autenticacao e permissoes em decoradores.
2. Separar acesso ao banco em funcoes/repositories.
3. Criar `base.html` para reduzir repeticao nos templates.
4. Mover CSS e JavaScript para arquivos estaticos.
5. Adicionar validacoes e protecoes nas acoes criticas.

