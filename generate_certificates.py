#!/usr/bin/env python3
"""
Automação para geração de certificados em PDF a partir de relatório de presença (CSV) e modelo PPTX.

Este script:
1. Lê um arquivo CSV exportado do Microsoft Teams com dados de presença
2. Extrai os nomes dos participantes
3. Para cada participante:
   - Cria uma cópia do template PPTX
   - Substitui o nome do participante
   - Converte para PDF usando PowerPoint COM
"""

import csv
import os
import shutil
import sys
from pathlib import Path
import re
import logging
from datetime import datetime
from dataclasses import dataclass
from typing import Callable, Optional

try:
    from pptx import Presentation
except ImportError:
    print("Erro: python-pptx não está instalado. Execute: pip install python-pptx")
    sys.exit(1)

try:
    from comtypes.client import CreateObject
except ImportError:
    print("Aviso: comtypes não está instalado. Execute: pip install comtypes")
    print("Será tentada apenas a conversão PPTX sem PDF.")
    CreateObject = None

# Caminhos base (usados pelo CLI e pelo logging)
BASE_DIR = Path(__file__).parent
LOGS_DIR = BASE_DIR / "logs"

# Cria diretório de logs antes de configurar logging
LOGS_DIR.mkdir(exist_ok=True)

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(str(LOGS_DIR / 'execution_log.txt')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


@dataclass
class CertificateConfig:
    csv_file: Path
    template_pptx: Path
    output_dir: Path
    temp_dir: Path
    progress_callback: Optional[Callable[[int, int, str], None]] = None


def create_directories(config: CertificateConfig):
    """Cria diretórios necessários."""
    config.output_dir.mkdir(exist_ok=True)
    config.temp_dir.mkdir(exist_ok=True)
    logger.info(f"Diretórios criados/verificados: {config.output_dir}, {config.temp_dir}")


def read_csv_participants(config: CertificateConfig):
    """
    Lê o arquivo CSV e extrai os nomes dos participantes.

    O CSV exportado do Teams tem a seguinte estrutura:
    - Linhas 0-8: Cabeçalho e seção de resumo
    - Linhas 9-42: Seção de Participantes (nome na coluna 0, função na coluna 6)
    - Depois: Seções de Atividades e Consentimento

    Returns:
        list: Lista de nomes únicos de participantes (strings)
    """
    participants = []
    seen_names = set()

    try:
        with open(config.csv_file, 'r', encoding='utf-16-le') as f:
            reader = csv.reader(f, delimiter='\t')

            # Pula as primeiras linhas até achar a seção "2. Participantes"
            in_participants_section = False
            for row_idx, row in enumerate(reader):
                if not row:
                    continue

                # Detecta início da seção de participantes
                if len(row) > 0 and "Participantes" in row[0] and "." in row[0]:
                    in_participants_section = True
                    logger.debug(f"Seção de Participantes encontrada na linha {row_idx}")
                    continue

                # Detecta fim da seção (começa nova seção com número)
                if in_participants_section and len(row) > 0 and re.match(r'^\d+\.\s+', row[0]):
                    logger.debug(f"Fim da seção de participantes na linha {row_idx}")
                    break

                # Se está na seção e a linha tem dados de participante
                if in_participants_section and len(row) >= 7:
                    name = row[0].strip() if row[0] else ""
                    function = row[6].strip() if len(row) > 6 and row[6] else ""

                    # Filtra nomes válidos (não vazios, sem serviços de bot)
                    if name and function in ["Participante", "Organizador"]:
                        # Remove sufixo "(Não verificado)"
                        clean_name = re.sub(r'\s*\(Não verificado\)\s*', '', name).strip()

                        # Evita duplicatas
                        if clean_name and clean_name not in seen_names:
                            participants.append(clean_name)
                            seen_names.add(clean_name)
                            logger.debug(f"Participante adicionado: {clean_name}")

    except FileNotFoundError:
        raise RuntimeError(f"Arquivo CSV não encontrado: {config.csv_file}")
    except Exception as e:
        raise RuntimeError(f"Erro ao ler CSV: {e}")

    logger.info(f"Total de {len(participants)} participantes extraídos do CSV")
    return participants


def sanitize_filename(name):
    """
    Sanitiza o nome para usar como nome de arquivo.

    Remove caracteres especiais e espaços.
    """
    # Remove caracteres inválidos em nomes de arquivo
    sanitized = re.sub(r'[<>:"/\\|?*]', '', name)
    sanitized = re.sub(r'\s+', '_', sanitized)
    return sanitized


def find_and_replace_name_in_pptx(pptx_path, participant_name):
    """
    Abre um arquivo PPTX e substitui o nome do participante.

    Procura por text boxes que contenham "Certificamos que" e substitui
    o nome imediatamente após esse texto.

    Args:
        pptx_path (Path): Caminho do arquivo PPTX
        participant_name (str): Nome do participante para inserir

    Returns:
        bool: True se a substituição foi bem-sucedida
    """
    try:
        prs = Presentation(str(pptx_path))

        # Itera sobre todos os slides (deve haver apenas 1)
        for slide in prs.slides:
            for shape in slide.shapes:
                # Verifica se é uma shape com texto
                if not hasattr(shape, "text"):
                    continue

                text = shape.text

                # Se encontrar o padrão "Certificamos que", substitui o nome depois
                if "Certificamos que" in text:
                    # Tenta encontrar e substituir no text_frame
                    if hasattr(shape, "text_frame"):
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                # Procura pelo placeholder ou padrão específico
                                if "[" in run.text or "NOME" in run.text or "nome" in run.text.lower():
                                    # Substitui placeholders comuns
                                    run.text = run.text.replace("[NOME DO PARTICIPANTE]", participant_name)
                                    run.text = run.text.replace("[NOME]", participant_name)
                                    run.text = run.text.replace("NOME DO PARTICIPANTE", participant_name)
                                    logger.debug(f"Placeholder substituído: {run.text}")

        # Se não encontrou nenhum placeholder com padrão, tenta uma abordagem mais simples
        # Procura por qualquer texto que possa ser um placeholder
        found_and_replaced = False
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            # Se o texto parece ser um placeholder (vazio, genérico, ou marcador)
                            if run.text.lower() in ["nome", "[nome]", "participante", "[participante]"]:
                                run.text = participant_name
                                found_and_replaced = True
                                logger.debug(f"Nome substituído em run simples: {participant_name}")

        # Salva o arquivo modificado
        prs.save(str(pptx_path))
        logger.info(f"PPTX modificado com sucesso: {pptx_path}")
        return True

    except Exception as e:
        logger.error(f"Erro ao modificar PPTX: {e}")
        return False


def convert_pptx_to_pdf(pptx_path, pdf_path):
    """
    Converte um arquivo PPTX para PDF usando PowerPoint COM.

    Args:
        pptx_path (Path): Caminho do arquivo PPTX de entrada
        pdf_path (Path): Caminho do arquivo PDF de saída

    Returns:
        bool: True se a conversão foi bem-sucedida
    """
    if CreateObject is None:
        logger.warning("comtypes não disponível - pulando conversão para PDF")
        logger.info(f"PPTX disponível em: {pptx_path}")
        return False

    try:
        # Abre o PowerPoint via COM
        ppt = CreateObject("PowerPoint.Application")
        ppt.Visible = True  # Deixa visível (workaround para o erro de "Hiding not allowed")

        # Abre a apresentação
        abs_pptx = os.path.abspath(str(pptx_path))
        abs_pdf = os.path.abspath(str(pdf_path))

        prs = ppt.Presentations.Open(abs_pptx)

        # Exporta para PDF (formato 32 = ppSaveAsPDF)
        prs.ExportAsFixedFormat(abs_pdf, 32)

        # Fecha a apresentação
        prs.Close()

        # Fecha o PowerPoint
        ppt.Quit()

        logger.info(f"PDF criado com sucesso: {pdf_path}")
        return True

    except Exception as e:
        logger.error(f"Erro ao converter PPTX para PDF: {e}")
        logger.warning(f"Arquivo PPTX mantido como fallback: {pptx_path}")
        return False


def generate_certificate(participant_name, config: CertificateConfig):
    """
    Gera um certificado personalizado para um participante.

    Processo:
    1. Copia o template PPTX para a pasta temporária
    2. Substitui o nome do participante
    3. Converte para PDF
    4. Move para pasta de saída

    Args:
        participant_name (str): Nome do participante
        config (CertificateConfig): Configuração com caminhos e callback

    Returns:
        bool: True se gerado com sucesso
    """
    try:
        # Cria nomes de arquivo
        safe_name = sanitize_filename(participant_name)
        temp_pptx = config.temp_dir / f"{safe_name}_temp.pptx"
        output_pdf = config.output_dir / f"Certificado_{safe_name}.pdf"

        # Copia o template para a pasta temporária
        shutil.copy2(config.template_pptx, temp_pptx)
        logger.debug(f"Template copiado para: {temp_pptx}")

        # Substitui o nome no PPTX
        if not find_and_replace_name_in_pptx(temp_pptx, participant_name):
            logger.error(f"Falha ao substituir nome no PPTX para: {participant_name}")
            return False

        # Tenta converter para PDF
        success = convert_pptx_to_pdf(temp_pptx, output_pdf)

        # Se a conversão para PDF falhou, mantém apenas o PPTX
        if not success:
            output_pptx = config.output_dir / f"Certificado_{safe_name}.pptx"
            shutil.move(temp_pptx, output_pptx)
            logger.warning(f"Certificado salvo como PPTX (falha na conversão PDF): {output_pptx}")
            return True

        # Remove o arquivo temporário PPTX se conversão foi bem-sucedida
        temp_pptx.unlink()

        return True

    except Exception as e:
        logger.error(f"Erro ao gerar certificado para {participant_name}: {e}")
        return False


def generate_all_certificates(participants, config: CertificateConfig):
    """
    Gera certificados para todos os participantes.

    Args:
        participants (list): Lista de nomes dos participantes
        config (CertificateConfig): Configuração com caminhos e callback

    Returns:
        dict: Relatório com estatísticas de sucesso/falha
    """
    total = len(participants)
    success = 0
    failed = []

    logger.info(f"Iniciando geração de {total} certificados...")

    for idx, participant in enumerate(participants, 1):
        logger.info(f"[{idx}/{total}] Gerando certificado para: {participant}")

        if generate_certificate(participant, config):
            success += 1
        else:
            failed.append(participant)

        if config.progress_callback:
            config.progress_callback(idx, total, participant)

        # Mostra progresso a cada 5 certificados
        if idx % 5 == 0 or idx == total:
            logger.info(f"Progresso: {success}/{idx} certificados gerados com sucesso")

    # Relatório final
    report = {
        'total': total,
        'success': success,
        'failed': len(failed),
        'failed_names': failed,
        'success_rate': (success / total * 100) if total > 0 else 0
    }

    return report


def print_report(report):
    """Imprime e registra o relatório final."""
    print("\n" + "=" * 60)
    print("RELATÓRIO DE GERAÇÃO DE CERTIFICADOS")
    print("=" * 60)
    print(f"Data/Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Total de participantes: {report['total']}")
    print(f"✓ Sucesso: {report['success']}")
    print(f"✗ Falhas: {report['failed']}")
    print(f"Taxa de sucesso: {report['success_rate']:.1f}%")

    if report['failed_names']:
        print("\nParticipantes com falha:")
        for name in report['failed_names']:
            print(f"  - {name}")

    print(f"\nCertificados salvos em: {report.get('output_dir', '')}")
    print("=" * 60 + "\n")

    # Registra no log
    logger.info("=" * 60)
    logger.info("RELATÓRIO FINAL DE GERAÇÃO")
    logger.info(f"Total: {report['total']}, Sucesso: {report['success']}, Falhas: {report['failed']}")
    logger.info(f"Taxa de sucesso: {report['success_rate']:.1f}%")
    logger.info("=" * 60)


def main():
    """Função principal (modo CLI)."""
    config = CertificateConfig(
        csv_file=BASE_DIR / "relatorio.csv",
        template_pptx=BASE_DIR / "modelo_certificado.pptx",
        output_dir=BASE_DIR / "certificados_gerados",
        temp_dir=BASE_DIR / ".temp_pptx",
    )

    logger.info("Iniciando automação de geração de certificados")
    logger.info(f"Arquivo CSV: {config.csv_file}")
    logger.info(f"Template PPTX: {config.template_pptx}")

    # Validações iniciais
    if not config.template_pptx.exists():
        logger.error(f"Template PPTX não encontrado: {config.template_pptx}")
        sys.exit(1)

    if not config.csv_file.exists():
        logger.error(f"Arquivo CSV não encontrado: {config.csv_file}")
        sys.exit(1)

    # Cria estrutura de diretórios
    create_directories(config)

    # Lê participantes do CSV
    participants = read_csv_participants(config)

    if not participants:
        logger.error("Nenhum participante encontrado no CSV")
        sys.exit(1)

    # Gera certificados
    report = generate_all_certificates(participants, config)

    # Imprime e registra o relatório
    print_report(report)

    # Retorna código apropriado
    if report['failed'] == 0:
        logger.info("Processamento concluído com sucesso!")
        return 0
    else:
        logger.warning(f"Processamento concluído com {report['failed']} falhas")
        return 1


if __name__ == "__main__":
    sys.exit(main())
