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
import shutil
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
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

import subprocess

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

SMTP_HOST = "10.181.30.1"
SMTP_PORT = 25
FROM_EMAIL = "automacao@mindworks.com.br"
ALLOWED_DOMAIN = "@mindworks.com.br"


@dataclass
class CertificateConfig:
    csv_file: Path
    template_pptx: Path
    output_dir: Path
    temp_dir: Path
    event_name: str = ""
    instructor_name: str = ""
    progress_callback: Optional[Callable[[int, int, str], None]] = None


def create_directories(config: CertificateConfig):
    """Cria diretórios necessários."""
    config.output_dir.mkdir(exist_ok=True)
    config.temp_dir.mkdir(exist_ok=True)
    logger.info(f"Diretórios criados/verificados: {config.output_dir}, {config.temp_dir}")


def read_csv_participants(config: CertificateConfig):
    """
    Lê o arquivo CSV (Nome;Email;Função) e extrai participantes.

    Returns:
        list[dict]: Lista de dicts com chaves 'name' e 'email'
    """
    participants = []
    seen_names = set()

    rows = None
    for encoding in ('utf-8-sig', 'cp1252', 'utf-8'):
        try:
            with open(config.csv_file, 'r', encoding=encoding) as f:
                rows = list(csv.reader(f, delimiter=';'))
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if rows is None:
        raise RuntimeError(f"Não foi possível ler o arquivo CSV: {config.csv_file}")

    for row_idx, row in enumerate(rows):
        if row_idx == 0:
            continue  # pula cabeçalho
        if not row or not row[0].strip():
            continue

        name = row[0].strip()
        email = row[1].strip() if len(row) > 1 else ""

        if name and name not in seen_names:
            participants.append({"name": name, "email": email})
            seen_names.add(name)
            logger.debug(f"Participante adicionado: {name} <{email}>")

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


def find_and_replace_in_pptx(pptx_path, replacements: dict):
    """
    Substitui todos os placeholders no PPTX conforme o dicionário fornecido.

    Args:
        pptx_path (Path): Caminho do arquivo PPTX
        replacements (dict): Mapa de placeholder → valor

    Returns:
        bool: True se bem-sucedido
    """
    try:
        prs = Presentation(str(pptx_path))
        for slide in prs.slides:
            for shape in slide.shapes:
                if not hasattr(shape, "text_frame"):
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for placeholder, value in replacements.items():
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, value)
                                logger.debug(f"Substituído '{placeholder}' → '{value}'")
        prs.save(str(pptx_path))
        logger.info(f"PPTX modificado com sucesso: {pptx_path}")
        return True
    except Exception as e:
        logger.error(f"Erro ao modificar PPTX: {e}")
        return False


def convert_pptx_to_pdf(pptx_path, pdf_path):
    """Converte PPTX para PDF usando LibreOffice headless."""
    try:
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf',
             '--outdir', str(pdf_path.parent), str(pptx_path)],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            logger.error(f"LibreOffice erro: {result.stderr}")
            return False
        # LibreOffice nomeia o PDF com o stem do PPTX de entrada
        generated = pdf_path.parent / (Path(pptx_path).stem + '.pdf')
        if generated.exists() and generated != pdf_path:
            generated.rename(pdf_path)
        logger.info(f"PDF criado: {pdf_path}")
        return True
    except Exception as e:
        logger.error(f"Erro ao converter PPTX para PDF: {e}")
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

        # Substitui placeholders no PPTX
        replacements = {
            "[NOME DO PARTICIPANTE]": participant_name,
            "[NOME]": participant_name,
            "NOME DO PARTICIPANTE": participant_name,
            "[EVENTO]": config.event_name,
            "[MINISTRANTE]": config.instructor_name,
        }
        if not find_and_replace_in_pptx(temp_pptx, replacements):
            logger.error(f"Falha ao substituir placeholders no PPTX para: {participant_name}")
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
        name = participant["name"] if isinstance(participant, dict) else participant
        logger.info(f"[{idx}/{total}] Gerando certificado para: {name}")

        if generate_certificate(name, config):
            success += 1
        else:
            failed.append(name)

        if config.progress_callback:
            config.progress_callback(idx, total, name)

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


def send_certificate_email(participant: dict, config: CertificateConfig) -> bool:
    """
    Envia o certificado PDF por email para um participante.

    Returns:
        bool: True se enviado com sucesso, False se ignorado ou com erro
    """
    name = participant["name"]
    email = participant.get("email", "")

    if not email.endswith(ALLOWED_DOMAIN):
        logger.warning(f"Email ignorado (domínio não autorizado): {name} <{email}>")
        return False

    pdf_path = config.output_dir / f"Certificado_{sanitize_filename(name)}.pdf"
    if not pdf_path.exists():
        logger.error(f"PDF não encontrado para envio: {pdf_path}")
        return False

    try:
        msg = MIMEMultipart()
        msg["From"] = FROM_EMAIL
        msg["To"] = email
        msg["Subject"] = f"Certificado - {config.event_name}"

        body = (
            f"Olá {name},\n\n"
            f"Segue em anexo seu certificado de participação no evento {config.event_name}.\n\n"
            f"Atenciosamente,\n"
            f"Mindworks"
        )
        msg.attach(MIMEText(body, "plain", "utf-8"))

        with open(pdf_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="Certificado_{sanitize_filename(name)}.pdf"'
            )
            msg.attach(part)

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as server:
            server.sendmail(FROM_EMAIL, email, msg.as_string())

        logger.info(f"Email enviado: {name} <{email}>")
        return True

    except Exception as e:
        logger.error(f"Erro ao enviar email para {name} <{email}>: {e}")
        return False


def send_all_emails(participants: list, config: CertificateConfig,
                    progress_callback: Optional[Callable[[int, int, str], None]] = None) -> dict:
    """
    Envia emails com certificados para todos os participantes.

    Returns:
        dict: Relatório com total, enviados, falhas e ignorados
    """
    total = len(participants)
    sent = 0
    failed = []
    skipped = []

    logger.info(f"Iniciando envio de {total} emails...")

    for idx, participant in enumerate(participants, 1):
        name = participant["name"]
        email = participant.get("email", "")

        if not email.endswith(ALLOWED_DOMAIN):
            logger.warning(f"[{idx}/{total}] Ignorado (domínio inválido): {name} <{email}>")
            skipped.append(name)
        elif send_certificate_email(participant, config):
            sent += 1
        else:
            failed.append(name)

        if progress_callback:
            progress_callback(idx, total, name)

    logger.info(f"Envio concluído: {sent} enviados, {len(failed)} falhas, {len(skipped)} ignorados")
    return {
        "total": total,
        "sent": sent,
        "failed": len(failed),
        "failed_names": failed,
        "skipped": len(skipped),
        "skipped_names": skipped,
    }


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
        csv_file=BASE_DIR / "participantes.CSV",
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
