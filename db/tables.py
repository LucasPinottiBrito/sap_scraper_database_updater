from sqlalchemy import Table, Column, String, MetaData

metadata = MetaData()

NotasAbertasTable = Table(
    "NotasAbertas",
    metadata,
    Column("NotaSAP", String, nullable=False),
    Column("TipoNota", String, nullable=False),
    Column("DataCriacao", String, nullable=False),
    Column("ConclusaoDesejada", String, nullable=False),
    Column("Estado", String, nullable=False),
    schema="Satisfacao"
)
