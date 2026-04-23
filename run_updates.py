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

from db.config import get_session
from db.models import Attachment, Note, TablesUpdatedAt
from db.tables import NotasAbertasTable
from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import datetime
from dotenv import load_dotenv
import os
from pathlib import Path
from typing import List, Dict
from sqlalchemy import delete, insert

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# Carregando credenciais de ambiente
SAP_USER = os.getenv("SAP_USER", "")
SAP_PASSWORD_EP1 = os.getenv("SAP_PASSWORD_EP1", "")
SAP_PASSWORD_EP2 = os.getenv("SAP_PASSWORD_EP2", "")

# Constantes
REGIONS = {
    "SP": {"ambiente": "EP1", "senha": SAP_PASSWORD_EP1},
    "ES": {"ambiente": "EP2", "senha": SAP_PASSWORD_EP2},
}

NOTE_TYPES = ["REC", "PRC", "SC/RC", "OVD"]


class UpdateAbortedError(Exception):
    """Erro usado para abortar a atualização antes de alterar a base."""


def close_screen(screen, label: str) -> None:
    if not screen:
        return
    try:
        screen.close()
    except Exception as e:
        print(f"  ⚠ Erro ao fechar {label}: {e}")


def touch_table_timestamp(session, table_name: str) -> None:
    tables_updated_at = session.get(TablesUpdatedAt, table_name)
    if not tables_updated_at:
        session.add(TablesUpdatedAt(table_name=table_name, updated_at=datetime.datetime.now()))
        return
    tables_updated_at.updated_at = datetime.datetime.now()


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

    iw67Screen = None
    all_notes = []
    successful_note_types = 0

    try:
        logon = SapLogonScreen()
        login = logon.loadSystem(regiao, ambiente)
        home = login.login(SAP_USER, senha, regiao, ambiente)
        iw67Screen = home.openTransaction("iw67")

        # Iterar sobre todos os tipos de notas
        for note_type in NOTE_TYPES:
            try:
                print(f"  📄 Buscando notas do tipo: {note_type}")
                notes_screen = iw67Screen.openNotesScreen(regiao, noteType=note_type)
                notes = notes_screen.getNotes(note_type)
                successful_note_types += 1
                all_notes.extend(notes)
                print(f"  ✓ {len(notes)} nota(s) encontrada(s) do tipo {note_type}")
            except Exception as e:
                print(f"  ⚠ Nenhuma nota encontrada para tipo {note_type}: {e}")
                continue
    finally:
        close_screen(iw67Screen, f"IW67 {regiao}")

    if successful_note_types == 0:
        raise UpdateAbortedError(f"Nenhum tipo de nota foi lido com sucesso na região {regiao}")

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

    iw52Screen = None
    enriched_notes = []

    try:
        logon = SapLogonScreen()
        login = logon.loadSystem(regiao, ambiente)
        home = login.login(SAP_USER, os.getenv(f"SAP_PASSWORD_{ambiente.upper()}"), regiao, ambiente)
        iw52Screen = home.openTransaction("iw52")

        for idx, note in enumerate(notes, 1):
            note_number = note.get("note_number")
            print(f"\n  [{idx}/{len(notes)}] Processando nota: {note_number}")

            iw52NoteScreen = None
            try:
                iw52NoteScreen = iw52Screen.openNote(note_number)

                # Determinar modelo da nota (CI = Recurso, NA = outros)
                note_model = "CI" if note.get("note_type") == "Recurso" else "NA"

                # Obter detalhes
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
                print(f"    ⚠ Erro ao enriquecer nota {note_number} ({regiao}): {e}")
                # raise UpdateAbortedError(f"Erro ao enriquecer nota {note_number} ({regiao}): {e}") from e
            finally:
                if iw52NoteScreen:
                    try:
                        iw52NoteScreen.back()
                    except Exception as e:
                        raise UpdateAbortedError(f"Erro ao voltar da nota {note_number} ({regiao}): {e}") from e
    finally:
        close_screen(iw52Screen, f"IW52 {regiao}")

    print(f"\n✓ {len(enriched_notes)} de {len(notes)} notas foram enriquecidas")
    return pd.DataFrame(enriched_notes)


def persist_notes_to_database(regiao: str, notas_df: pd.DataFrame, session) -> None:
    """
    Persiste as notas enriquecidas no banco de dados.

    Para cada nota, salva:
    - Informações básicas na tabela Note
    - Anexos associados na tabela Attachment

    Notas que já existem no banco são atualizadas; novas notas são criadas.

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
    updated_count = 0

    for idx, nota_dict in enumerate(notas_list, 1):
        note_number = nota_dict.get("note_number")
        print(f"\n  [{idx}/{len(notas_list)}] Processando: {note_number}")

        try:
            payload = {
                "note_number": nota_dict.get("note_number"),
                "created_at": nota_dict.get("created_at"),
                "conclusion_date": nota_dict.get("conclusion_date"),
                "priority_text": nota_dict.get("priority_text"),
                "group_": nota_dict.get("note_type"),
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

            nota_existente = session.query(Note).filter_by(note_number=note_number).one_or_none()
            if nota_existente:
                for key, value in payload.items():
                    setattr(nota_existente, key, value)
                nota = nota_existente
                updated_count += 1
                print(f"    ↻ Nota atualizada")
            else:
                nota = Note(**payload)
                session.add(nota)
                created_count += 1
                print(f"    ✓ Nota salva com sucesso")

            session.flush()
            touch_table_timestamp(session, "sapAutoTPendingNotes")

            # Sincronizar anexos
            attachments = nota_dict.get("attachments", [])
            session.query(Attachment).filter_by(note_id=nota.id).delete(synchronize_session=False)
            for attachment_url in attachments:
                session.add(
                    Attachment(
                        note_id=nota.id,
                        url=attachment_url,
                        created_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    )
                )
            touch_table_timestamp(session, "sapAutoTAttachments")
            print(f"      📎 {len(attachments)} anexo(s) sincronizado(s)")

        except Exception as e:
            raise UpdateAbortedError(f"Erro ao salvar nota {note_number}: {e}") from e

    print(f"\n✓ Resultado: {created_count} criadas, {updated_count} atualizadas")


def update_bi_table(notas_df: pd.DataFrame, session) -> None:
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

    required_columns = ["TipoNota", "NotaSAP", "DataCriacao", "ConclusaoDesejada", "Estado"]
    if bi_notes.empty:
        bi_notes = pd.DataFrame(columns=required_columns)
    else:
        missing_columns = [column for column in required_columns if column not in bi_notes.columns]
        if missing_columns:
            raise UpdateAbortedError(f"Colunas obrigatórias ausentes para BI: {missing_columns}")
        bi_notes = bi_notes[required_columns]

    # Converter para dicionários e fazer replace
    replace_records = bi_notes.to_dict(orient="records")
    
    print(f"  Substituindo {len(replace_records)} registros na BI...")
    session.execute(delete(NotasAbertasTable))
    if replace_records:
        session.execute(insert(NotasAbertasTable), replace_records)
    touch_table_timestamp(session, "NotasAbertasTable")
    
    print(f"  ✓ Tabela BI atualizada com sucesso")


def clean_old_notes(all_note_numbers: List[str], session) -> None:
    """
    Remove notas antigas do banco de dados que não estão mais na lista atual.

    Mantém apenas as notas que estão no sistema SAP (as buscadas recentemente).

    Args:
        all_note_numbers (List[str]): Lista de números de notas para manter
    """
    print(f"\n{'='*60}")
    print(f"🧹 Limpando notas antigas do banco de dados")
    print(f"{'='*60}")

    if all_note_numbers:
        deleted_count = session.query(Note).filter(~Note.note_number.in_(all_note_numbers)).delete(synchronize_session=False)
    else:
        deleted_count = session.query(Note).delete(synchronize_session=False)
    touch_table_timestamp(session, "sapAutoTPendingNotes")
    print(f"  ✓ {deleted_count} nota(s) removida(s)")


def _run_full_update_unlocked() -> None:
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
                raise UpdateAbortedError(f"Erro ao buscar notas da região {regiao}: {e}") from e

        # 2. ENRIQUECER TODAS AS REGIÕES ANTES DE ALTERAR A BASE
        for regiao, data in all_regions_data.items():
            notas_basicas = data["notas_basicas"]
            ambiente = data["ambiente"]

            if notas_basicas.empty:
                print(f"\n⚠ Pulando {regiao}: sem notas para processar")
                all_regions_data[regiao]["notas_finais"] = pd.DataFrame()
                continue

            try:
                # Enriquecer notas
                notas_enriquecidas = enrich_notes_with_details(
                    ambiente,
                    regiao,
                    notas_basicas.to_dict(orient="records"),
                )

                # Armazenar para BI
                notas_enriquecidas["region"] = regiao
                all_regions_data[regiao]["notas_finais"] = notas_enriquecidas

            except Exception as e:
                raise UpdateAbortedError(f"Erro ao processar região {regiao}: {e}") from e

        # 3. PREPARAR DADOS DA BI
        notes_for_bi = [
            data.get("notas_finais", pd.DataFrame())
            for data in all_regions_data.values()
            if not data.get("notas_finais", pd.DataFrame()).empty
        ]
        all_notes_for_bi = pd.concat(notes_for_bi, ignore_index=True) if notes_for_bi else pd.DataFrame()

        # 4. SOMENTE AGORA ALTERAR A BASE, EM UMA ÚNICA TRANSAÇÃO
        session = get_session()
        try:
            with session.begin():
                for regiao, data in all_regions_data.items():
                    persist_notes_to_database(regiao, data.get("notas_finais", pd.DataFrame()), session)

                update_bi_table(all_notes_for_bi, session)
                clean_old_notes(all_note_numbers, session)
        finally:
            session.close()

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


def run_full_update(acquire_lock: bool = True) -> None:
    if not acquire_lock:
        _run_full_update_unlocked()
        return

    from sap_execution_lock import SapExecutionLock, SapExecutionLockBusy

    try:
        with SapExecutionLock(operation="atualizacao de bases SAP"):
            _run_full_update_unlocked()
    except SapExecutionLockBusy as exc:
        raise UpdateAbortedError("SAP ocupado por outra execucao. Atualizacao nao iniciada.") from exc


if __name__ == "__main__":
    run_full_update()
