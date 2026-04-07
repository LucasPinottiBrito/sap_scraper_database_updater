"""
Teste e Validação do Action Queue Processor

Este script testa:
1. Importação dos módulos
2. Estrutura das classes
3. Conexão com banco de dados
4. Fluxo básico sem executar ações reais
"""

import sys
import logging
from datetime import datetime

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


def test_imports():
    """Testa se todos os imports funcionam"""
    logger.info("=" * 60)
    logger.info("TESTE 1: Validando Imports")
    logger.info("=" * 60)
    
    try:
        from action_processor import (
            ActionExecutor, 
            ActionQueueProcessor, 
            ExecutionResult,
            run_queue_processor
        )
        logger.info("✓ action_processor importado com sucesso")
        
        from db.repositories.SAPActionQueue import SAPActionQueueRepository
        logger.info("✓ SAPActionQueueRepository importado com sucesso")
        
        from db.repositories.Note import NoteRepository
        logger.info("✓ NoteRepository importado com sucesso")
        
        from db.models import SAPActionQueue
        logger.info("✓ SAPActionQueue model importado com sucesso")
        
        logger.info("\n✓ TODOS OS IMPORTS OK\n")
        return True
        
    except Exception as e:
        logger.error(f"\n✗ ERRO DE IMPORT: {str(e)}\n")
        return False


def test_action_handlers():
    """Testa se os handlers de ação estão registrados"""
    logger.info("=" * 60)
    logger.info("TESTE 2: Validando Handlers de Ação")
    logger.info("=" * 60)
    
    try:
        from action_processor import ActionQueueProcessor
        
        processor = ActionQueueProcessor()
        handlers = processor.ACTION_HANDLERS
        
        required_actions = [
            "proceder",
            "improceder",
            "transferir_semci",
            "transferir_mmgd",
            "encerrar"
        ]
        
        for action in required_actions:
            if action in handlers:
                logger.info(f"✓ Handler '{action}' registrado")
            else:
                logger.error(f"✗ Handler '{action}' NÃO encontrado")
                return False
        
        logger.info(f"\n✓ TODOS OS {len(required_actions)} HANDLERS OK\n")
        return True
        
    except Exception as e:
        logger.error(f"\n✗ ERRO: {str(e)}\n")
        return False


def test_database_connection():
    """Testa conexão com banco de dados"""
    logger.info("=" * 60)
    logger.info("TESTE 3: Validando Conexão com Banco de Dados")
    logger.info("=" * 60)
    
    try:
        from db.repositories.SAPActionQueue import SAPActionQueueRepository
        
        repo = SAPActionQueueRepository()
        logger.info("✓ SAPActionQueueRepository inicializado")
        
        # Tentar fazer uma query simples
        pending = repo.list_pending()
        logger.info(f"✓ Conexão ativa - Encontradas {len(pending)} ações pendentes")
        
        logger.info("\n✓ BANCO DE DADOS OK\n")
        return True
        
    except Exception as e:
        logger.error(f"\n✗ ERRO DE CONEXÃO: {str(e)}\n")
        return False


def test_processor_structure():
    """Testa estrutura do processador"""
    logger.info("=" * 60)
    logger.info("TESTE 4: Validando Estrutura do Processador")
    logger.info("=" * 60)
    
    try:
        from action_processor import (
            ActionQueueProcessor,
            ExecutionResult
        )
        
        processor = ActionQueueProcessor()
        logger.info("✓ ActionQueueProcessor inicializado")
        
        # Verificar métodos principais
        methods = [
            'fetch_pending_actions',
            'validate_action',
            'build_execution_context',
            'execute_action',
            'update_action_status',
            'delete_related_note',
            'process_queue'
        ]
        
        for method in methods:
            if hasattr(processor, method):
                logger.info(f"✓ Método '{method}' existe")
            else:
                logger.error(f"✗ Método '{method}' NÃO encontrado")
                return False
        
        logger.info("\n✓ ESTRUTURA DO PROCESSADOR OK\n")
        return True
        
    except Exception as e:
        logger.error(f"\n✗ ERRO: {str(e)}\n")
        return False


def test_execution_result():
    """Testa classe ExecutionResult"""
    logger.info("=" * 60)
    logger.info("TESTE 5: Validando ExecutionResult")
    logger.info("=" * 60)
    
    try:
        from action_processor import ExecutionResult
        
        # Teste com sucesso
        result_ok = ExecutionResult(action_id=1, success=True)
        logger.info(f"✓ ExecutionResult (sucesso): {result_ok.to_dict()}")
        
        # Teste com erro
        result_error = ExecutionResult(
            action_id=2, 
            success=False, 
            error_message="Erro de teste"
        )
        logger.info(f"✓ ExecutionResult (erro): {result_error.to_dict()}")
        
        logger.info("\n✓ EXECUTIONRESULT OK\n")
        return True
        
    except Exception as e:
        logger.error(f"\n✗ ERRO: {str(e)}\n")
        return False


def main():
    """Executa todos os testes"""
    logger.info("\n")
    logger.info("╔" + "=" * 58 + "╗")
    logger.info("║" + " " * 58 + "║")
    logger.info("║" + "  TESTE DO ACTION QUEUE PROCESSOR".center(58) + "║")
    logger.info("║" + " " * 58 + "║")
    logger.info("╚" + "=" * 58 + "╝")
    logger.info("\n")
    
    tests = [
        ("Imports", test_imports),
        ("Action Handlers", test_action_handlers),
        ("Database Connection", test_database_connection),
        ("Processor Structure", test_processor_structure),
        ("ExecutionResult", test_execution_result),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            logger.error(f"EXCEÇÃO NÃO TRATADA em {test_name}: {str(e)}\n")
            results.append((test_name, False))
    
    # Resumo
    logger.info("=" * 60)
    logger.info("RESUMO DOS TESTES")
    logger.info("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✓ PASSOU" if result else "✗ FALHOU"
        logger.info(f"{status}: {test_name}")
    
    logger.info("\n" + "=" * 60)
    logger.info(f"RESULTADO: {passed}/{total} testes passaram")
    logger.info("=" * 60 + "\n")
    
    if passed == total:
        logger.info("🎉 TODOS OS TESTES PASSARAM - SISTEMA PRONTO\n")
        return 0
    else:
        logger.info("❌ ALGUNS TESTES FALHARAM - VERIFIQUE ACIMA\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
