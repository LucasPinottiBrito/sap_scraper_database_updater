"""
Módulo unificado para atualização de dados SAP.

Este módulo centraliza toda a lógica de busca de notas no SAP, enriquecimento
com detalhes e anexos, persistência no banco de dados e atualização da tabela BI.

A principal vantagem é que as notas são buscadas apenas uma vez por região,
evitando múltiplas requisições ao SAP desnecessárias.

Fluxo:
1. Buscar notas do SAP (1x por região)
2. Enriquecer com detalhes e anexos (IW52)
3. Persistir no banco de dados
4. Atualizar tabela BI com dados agregados
"""

from db.repositories.TablesUpdatedAt import TablesUpdatedAtRepository
from db.repositories.Note import NoteRepository
from db.repositories.BIEntity import BIEntityRepository
from db.repositories.Attachment import AttachmentRepository
from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import datetime
from dotenv import load_dotenv
import os
from typing import List, Dict, Tuple

load_dotenv()

# Carregando credenciais de ambiente
SAP_USER = os.getenv("SAP_USER", "")
SAP_PASSWORD_EP1 = os.getenv("SAP_PASSWORD_EP1", "")
SAP_PASSWORD_EP2 = os.getenv("SAP_PASSWORD_EP2", "")

# Repositórios
updated_at_repo = TablesUpdatedAtRepository()
note_repo = NoteRepository()
bi_entity_repo = BIEntityRepository()
attach_repo = AttachmentRepository()

# Constantes
REGIONS = {
    "SP": {"ambiente": "EP1", "senha": SAP_PASSWORD_EP1},
    "ES": {"ambiente": "EP2", "senha": SAP_PASSWORD_EP2},
}

NOTE_TYPES = ["REC", "PRC", "SC/RC", "OVD"]


def fetch_notes_from_sap(ambiente: str, regiao: str, senha: str) -> pd.DataFrame:
    """
    Busca notas do SAP de uma determinada região e ambiente.

    Obtém notas de todos os tipos (REC, PRC, SC/RC, OVD) via transação IW67.
    Se algum tipo de nota não for encontrado, continua com os outros tipos.

    Args:
        ambiente (str): Ambiente SAP (ex: "EP1", "EP2")
        regiao (str): Região (ex: "SP", "ES")
        senha (str): Senha para acesso ao SAP

    Returns:
        pd.DataFrame: DataFrame com todas as notas encontradas

    Raises:
        Exception: Se não conseguir acessar o SAP ou abrir a transação IW67
    """
    print(f"\n{'='*60}")
    print(f"📌 Iniciando busca de notas - Região: {regiao}, Ambiente: {ambiente}")
    print(f"{'='*60}")

    logon = SapLogonScreen()
    login = logon.loadSystem(regiao, ambiente)
    home = login.login(SAP_USER, senha, regiao, ambiente)
    iw67Screen = home.openTransaction("iw67")

    all_notes = []

    # Iterar sobre todos os tipos de notas
    for note_type in NOTE_TYPES:
        try:
            print(f"  📄 Buscando notas do tipo: {note_type}")
            notes_screen = iw67Screen.openNotesScreen(regiao, noteType=note_type)
            notes = notes_screen.getNotes(note_type)
            all_notes.extend(notes)
            print(f"  ✓ {len(notes)} nota(s) encontrada(s) do tipo {note_type}")
        except Exception as e:
            print(f"  ⚠ Nenhuma nota encontrada para tipo {note_type}: {e}")
            continue

    notes_screen.close()
    df = pd.DataFrame(all_notes)

    print(f"\n✓ Total de {len(all_notes)} notas buscadas da região {regiao}")
    return df


def enrich_notes_with_details(
    ambiente: str, regiao: str, notes: List[Dict]
) -> pd.DataFrame:
    """
    Enriquece as notas com detalhes e informações de anexos via transação IW52.

    Para cada nota, obtém:
    - Descrição detalhada
    - Informações de contato (email, SMS, código)
    - Instituição
    - Anexos

    Args:
        ambiente (str): Ambiente SAP
        regiao (str): Região
        notes (List[Dict]): Lista de notas a enriquecer

    Returns:
        pd.DataFrame: DataFrame com notas enriquecidas
    """
    print(f"\n{'='*60}")
    print(f"📋 Iniciando enriquecimento de notas - {regiao}")
    print(f"{'='*60}")

    if not notes:
        print("  ⚠ Nenhuma nota para enriquecer")
        return pd.DataFrame()

    logon = SapLogonScreen()
    login = logon.loadSystem(regiao, ambiente)
    home = login.login(SAP_USER, os.getenv(f"SAP_PASSWORD_{ambiente.upper()}"), regiao, ambiente)
    iw52Screen = home.openTransaction("iw52")

    enriched_notes = []

    for idx, note in enumerate(notes, 1):
        note_number = note.get("note_number")
        print(f"\n  [{idx}/{len(notes)}] Processando nota: {note_number}")

        try:
            iw52NoteScreen = iw52Screen.openNote(note_number)
            
            # Determinar modelo da nota (CI = Recurso, NA = outros)
            note_model = "CI" if note.get("note_type") == "Recurso" else "NA"
            
            # Obter detalhes
            try:
                details = iw52NoteScreen.getNoteDetails(note_model=note_model)
                
                # Adicionar informações de detalhes
                note["description_detail"] = details.get("description_text", "")
                note["email"] = details.get("contato_email", "")
                note["sms"] = details.get("contato_sms", "")
                note["cod_contact"] = details.get("cod_contato", "")
                note["inst"] = details.get("inst", "")
                note["nome_cliente"] = details.get("nome_cliente", "")
                
                # Obter anexos
                attachments = iw52NoteScreen.get_attachments(
                    download_files=False,
                    folder_path_to_download=f"./attachments/{ambiente}/{note_number}/",
                )
                note["attachments"] = attachments or []
                
                enriched_notes.append(note)
                print(f"    ✓ Nota enriquecida com {len(note['attachments'])} anexo(s)")

            except Exception as e:
                note["description_detail"] = ""
                note["email"] = ""
                note["sms"] = ""
                note["cod_contact"] = ""
                note["inst"] = ""
                note["nome_cliente"] = ""
                note["attachments"] = []
                enriched_notes.append(note)
                print(f"    ⚠ Erro ao enriquecer nota {note_number}: {e}")

            iw52NoteScreen.back()

        except Exception as e:
            print(f"    ✗ Erro ao processar nota {note_number}: {e}")
            continue

    iw52Screen.close()

    print(f"\n✓ {len(enriched_notes)} de {len(notes)} notas foram enriquecidas")
    return pd.DataFrame(enriched_notes)


def persist_notes_to_database(regiao: str, notas_df: pd.DataFrame) -> None:
    """
    Persiste as notas enriquecidas no banco de dados.

    Para cada nota, salva:
    - Informações básicas na tabela Note
    - Anexos associados na tabela Attachment

    Notas que já existem no banco são puladas (evita duplicatas).

    Args:
        regiao (str): Região das notas
        notas_df (pd.DataFrame): DataFrame com notas a persistir
    """
    print(f"\n{'='*60}")
    print(f"💾 Persistindo notas no banco de dados - {regiao}")
    print(f"{'='*60}")

    if notas_df.empty:
        print("  ⚠ Nenhuma nota para persistir")
        return

    notas_list = notas_df.to_dict(orient="records")

    created_count = 0
    skipped_count = 0

    for idx, nota_dict in enumerate(notas_list, 1):
        note_number = nota_dict.get("note_number")
        print(f"\n  [{idx}/{len(notas_list)}] Processando: {note_number}")

        # Verificar se nota já existe
        nota_existente = note_repo.get_from_number(note_number)
        if nota_existente:
            print(f"    ⊘ Nota já existe no banco. Pulando...")
            skipped_count += 1
            continue
        

        try:
            # Criar nota no banco
            nota = note_repo.create(
                {
                    "note_number": nota_dict.get("note_number"),
                    "created_at": nota_dict.get("created_at"),
                    "conclusion_date": nota_dict.get("conclusion_date"),
                    "priority_text": nota_dict.get("priority_text"),
                    "group": nota_dict.get("note_type"),
                    "code_text": nota_dict.get("code_text"),
                    "city": nota_dict.get("city"),
                    "description": nota_dict.get("description"),
                    "description_detail": nota_dict.get("description_detail", ""),
                    "business_partner_id": nota_dict.get("business_partner_id"),
                    "email": nota_dict.get("email", ""),
                    "sms": nota_dict.get("sms", ""),
                    "cod_contact": nota_dict.get("cod_contact", ""),
                    "inst": nota_dict.get("inst", ""),
                    "region": regiao,
                    "nome_cliente": nota_dict.get("nome_cliente", ""),
                }
            )

            updated_at_repo.update_table_timestamp("sapAutoTPendingNotes")

            # Adicionar anexos
            attachments = nota_dict.get("attachments", [])
            if attachments:
                for attachment_url in attachments:
                    try:
                        attach_repo.create(
                            note_id=nota.id,
                            url=attachment_url,
                            created_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        )
                        updated_at_repo.update_table_timestamp("sapAutoTAttachments")
                        print(f"      📎 Anexo adicionado: {attachment_url}")
                    except Exception as e:
                        print(f"      ✗ Erro ao adicionar anexo {attachment_url}: {e}")

            created_count += 1
            print(f"    ✓ Nota salva com sucesso")

        except Exception as e:
            print(f"    ✗ Erro ao salvar nota {note_number}: {e}")
            continue

    print(f"\n✓ Resultado: {created_count} criadas, {skipped_count} puladas")


def update_bi_table(notas_df: pd.DataFrame) -> None:
    """
    Atualiza a tabela BI com os dados das notas.

    Renomeia as colunas para os nomes esperados pela tabela BI e
    substitui todos os registros.

    Args:
        notas_df (pd.DataFrame): DataFrame com notas a atualizar na BI
    """
    print(f"\n{'='*60}")
    print(f"📊 Atualizando tabela BI")
    print(f"{'='*60}")

    if notas_df.empty:
        print("  ⚠ Nenhuma nota para atualizar na BI")
        return

    # Renomear colunas para o padrão esperado pela BI
    bi_notes = notas_df.copy()
    bi_notes.rename(
        columns={
            "note_type": "TipoNota",
            "note_number": "NotaSAP",
            "created_at": "DataCriacao",
            "conclusion_date": "ConclusaoDesejada",
            "region": "Estado",
        },
        inplace=True,
    )

    # Converter para dicionários e fazer replace
    replace_records = bi_notes.to_dict(orient="records")
    
    print(f"  Substituindo {len(replace_records)} registros na BI...")
    bi_entity_repo.replace_all(replace_records)
    updated_at_repo.update_table_timestamp("NotasAbertasTable")
    
    print(f"  ✓ Tabela BI atualizada com sucesso")


def clean_old_notes(all_note_numbers: List[str]) -> None:
    """
    Remove notas antigas do banco de dados que não estão mais na lista atual.

    Mantém apenas as notas que estão no sistema SAP (as buscadas recentemente).

    Args:
        all_note_numbers (List[str]): Lista de números de notas para manter
    """
    print(f"\n{'='*60}")
    print(f"🧹 Limpando notas antigas do banco de dados")
    print(f"{'='*60}")

    try:
        deleted_count = note_repo.delete_all(notes_to_keep=all_note_numbers)
        print(f"  ✓ {deleted_count} nota(s) removida(s)")
    except Exception as e:
        print(f"  ✗ Erro ao limpar notas antigas: {e}")


def run_full_update() -> None:
    """
    Orquestra o fluxo completo de atualização de dados SAP.

    Executa sequencialmente:
    1. Busca de notas do SAP (uma vez por região)
    2. Enriquecimento com detalhes e anexos
    3. Persistência no banco de dados
    4. Atualização da tabela BI
    5. Limpeza de notas antigas

    Este é o ponto de entrada principal do módulo.
    """
    print("\n" + "="*60)
    print("🚀 INICIANDO PROCESSO DE ATUALIZAÇÃO COMPLETA DE DADOS SAP")
    print("="*60)

    start_time = datetime.datetime.now()
    all_regions_data = {}
    all_note_numbers = []

    try:
        # 1. BUSCAR NOTAS DO SAP (uma vez por região)
        for regiao, config in REGIONS.items():
            try:
                notes_df = fetch_notes_from_sap(
                    config["ambiente"],
                    regiao,
                    config["senha"],
                )
                all_regions_data[regiao] = {
                    "notas_basicas": notes_df,
                    "ambiente": config["ambiente"],
                }
                all_note_numbers.extend(notes_df.get("note_number", []).tolist())
            except Exception as e:
                print(f"\n✗ Erro ao buscar notas da região {regiao}: {e}")
                all_regions_data[regiao] = {"notas_basicas": pd.DataFrame(), "ambiente": config["ambiente"]}

        # 2. ENRIQUECER E PERSISTIR POR REGIÃO
        for regiao, data in all_regions_data.items():
            notas_basicas = data["notas_basicas"]
            ambiente = data["ambiente"]

            if notas_basicas.empty:
                print(f"\n⚠ Pulando {regiao}: sem notas para processar")
                continue

            try:
                # Enriquecer notas
                notas_enriquecidas = enrich_notes_with_details(
                    ambiente,
                    regiao,
                    notas_basicas.to_dict(orient="records"),
                )

                # Persistir no banco
                persist_notes_to_database(regiao, notas_enriquecidas)

                # Armazenar para BI
                notas_enriquecidas["region"] = regiao
                all_regions_data[regiao]["notas_finais"] = notas_enriquecidas

            except Exception as e:
                print(f"\n✗ Erro ao processar região {regiao}: {e}")
                all_regions_data[regiao]["notas_finais"] = pd.DataFrame()

        # 3. ATUALIZAR BI COM TODAS AS NOTAS
        all_notes_for_bi = pd.concat(
            [
                data.get("notas_finais", pd.DataFrame())
                for data in all_regions_data.values()
                if not data.get("notas_finais", pd.DataFrame()).empty
            ],
            ignore_index=True,
        )

        if not all_notes_for_bi.empty:
            update_bi_table(all_notes_for_bi)
        else:
            print("\n⚠ Nenhuma nota para atualizar na BI")

        # 4. LIMPAR NOTAS ANTIGAS
        if all_note_numbers:
            clean_old_notes(all_note_numbers)

        # Sucesso!
        elapsed_time = datetime.datetime.now() - start_time
        print(f"\n" + "="*60)
        print(f"✅ PROCESSO COMPLETADO COM SUCESSO")
        print(f"   Tempo total: {elapsed_time}")
        print("="*60 + "\n")

    except Exception as e:
        elapsed_time = datetime.datetime.now() - start_time
        print(f"\n" + "="*60)
        print(f"❌ ERRO NO PROCESSO: {e}")
        print(f"   Tempo até erro: {elapsed_time}")
        print("="*60 + "\n")
        raise


if __name__ == "__main__":
    run_full_update()
