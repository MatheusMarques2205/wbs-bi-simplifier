# WBS BI Simplifier

## Sobre
Ferramenta para simplificar dados de WBS do MS Project para visualização em Power BI, focada em projetos de energia eólica. Transforma milhares de tarefas em 19 marcos principais.

## Como Usar

### Requisitos
- Python 3.x
- Microsoft Access
- Power BI Desktop
- Pacotes Python: pyodbc, pandas, sqlalchemy

### Passo a Passo

1. Configure o arquivo Python:
   - Abra "WBS (Work Breakdown Structure).py"
   - Atualize o caminho do banco de dados:
   ```python
   DATABASE_PATH = "seu_caminho_do_banco.mdb"
   ```

2. Execute o programa:
   - Abra o terminal
   - Execute: `python "WBS (Work Breakdown Structure).py"`

3. No Power BI:
   - Importe as duas tabelas usando Power Query
   - Crie relacionamento entre as tabelas usando TASK_ID
   - Crie suas visualizações

## Marcos Principais
O programa categoriza as tarefas em:
- NTP
- Anchor Cages (Ex Works, POD, Site)
- WTG Components
- Foundation
- Crane Operations
- Installation
- Commissioning
- COD

## Suporte
Para dúvidas ou problemas, abra uma issue no GitHub.
