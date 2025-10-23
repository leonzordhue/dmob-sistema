import csv
import sys
import os
from datetime import datetime
from werkzeug.security import generate_password_hash

# Adiciona o diretório atual ao path para importar o app
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from app import app, db, Usuario, Ramal
    print("✅ Módulos importados com sucesso!")
except ImportError as e:
    print(f"❌ Erro na importação: {e}")
    print("💡 Certifique-se de que o app.py está na mesma pasta")
    sys.exit(1)

def limpar_ramais_existentes():
    """Remove todos os ramais existentes para importação limpa"""
    with app.app_context():
        try:
            num_ramais = Ramal.query.count()
            Ramal.query.delete()
            db.session.commit()
            print(f"🗑️  {num_ramais} ramais antigos removidos")
            return True
        except Exception as e:
            print(f"❌ Erro ao limpar ramais: {e}")
            db.session.rollback()
            return False

def importar_dados_csv():
    """Importa dados do CSV para o banco de dados - VERSÃO CORRIGIDA"""
    with app.app_context():
        print("🚀 INICIANDO IMPORTAÇÃO CORRIGIDA...")
        
        # Primeiro limpa os ramais existentes
        if not limpar_ramais_existentes():
            print("❌ Não foi possível limpar os ramais existentes")
            return
        
        # Contadores
        ramais_importados = 0
        erros = 0
        
        arquivo_csv = 'PLANILHA GERAL DEPARTAMENTO RODOVIÁRIO.csv'
        
        if not os.path.exists(arquivo_csv):
            print(f"❌ Arquivo CSV não encontrado: {arquivo_csv}")
            return
        
        try:
            with open(arquivo_csv, 'r', encoding='latin-1') as file:
                reader = csv.DictReader(file, delimiter=';')
                
                print(f"📖 Lendo dados do CSV...")
                
                # COLUNAS CORRETAS (baseado no debug anterior)
                coluna_codigo = ' CÓDIGO '
                coluna_numero = 'Número'
                coluna_ramal_estrada = 'Ramal/Estrada'  # COLUNA CORRETA PARA O NOME
                coluna_municipio = 'Município'
                coluna_extensao = 'Extensão (km)'
                coluna_situacao = 'Situação'
                coluna_revestimento = 'Revestimento'
                coluna_inicio = 'Local de Início '
                coluna_fim = 'Local Termino '
                coluna_lat_inicial = 'Latitude Inicial'
                coluna_lon_inicial = 'Longitude Inicial'
                coluna_lat_final = 'Latitude Final'
                coluna_lon_final = 'Longitude Final'
                
                for linha_num, linha in enumerate(reader, 2):
                    # Pula linhas vazias
                    if not linha.get(coluna_codigo) or not linha.get(coluna_ramal_estrada):
                        continue
                    
                    try:
                        # Dados obrigatórios
                        codigo = linha[coluna_codigo].strip()
                        ramal_estrada = linha[coluna_ramal_estrada].strip()  # NOME CORRETO DO RAMAL
                        
                        # Dados opcionais
                        numero_str = linha.get(coluna_numero, '').strip()
                        numero = int(numero_str) if numero_str and numero_str.isdigit() else None
                        
                        municipio = linha.get(coluna_municipio, '').strip() or None
                        
                        # Extensão
                        extensao_str = linha.get(coluna_extensao, '0').strip().replace(',', '.')
                        extensao_km = 0.0
                        try:
                            extensao_clean = ''.join(c for c in extensao_str if c.isdigit() or c == '.')
                            extensao_km = float(extensao_clean) if extensao_clean else 0.0
                        except (ValueError, TypeError):
                            extensao_km = 0.0
                        
                        # Situação e revestimento
                        situacao = linha.get(coluna_situacao, '').strip() or None
                        revestimento = linha.get(coluna_revestimento, '').strip() or None
                        
                        # Localização
                        inicio = linha.get(coluna_inicio, '').strip() or None
                        fim = linha.get(coluna_fim, '').strip() or None
                        
                        # Coordenadas
                        lat_inicial = linha.get(coluna_lat_inicial, '').strip() or None
                        lon_inicial = linha.get(coluna_lon_inicial, '').strip() or None
                        lat_final = linha.get(coluna_lat_final, '').strip() or None
                        lon_final = linha.get(coluna_lon_final, '').strip() or None
                        
                        coordenada_inicio = f"{lat_inicial}, {lon_inicial}" if lat_inicial and lon_inicial else None
                        coordenada_fim = f"{lat_final}, {lon_final}" if lat_final and lon_final else None
                        
                        # Para debug, mostra os primeiros 5 registros
                        if ramais_importados < 5:
                            print(f"   📥 Importando: {codigo} - {ramal_estrada} - {municipio}")
                        
                        # Cria o ramal
                        ramal = Ramal(
                            codigo=codigo,
                            numero=numero,
                            ramal_estrada=ramal_estrada,  # NOME CORRETO
                            municipio=municipio,
                            extensao_km=extensao_km,
                            situacao=situacao,
                            revestimento=revestimento,
                            inicio=inicio,
                            fim=fim,
                            coordenada_inicio=coordenada_inicio,
                            coordenada_fim=coordenada_fim
                        )
                        
                        db.session.add(ramal)
                        ramais_importados += 1
                        
                        # Mostra progresso a cada 50 registros
                        if ramais_importados % 50 == 0:
                            print(f"   📦 Processados: {ramais_importados} ramais...")
                        
                        # Commit a cada 100 registros
                        if ramais_importados % 100 == 0:
                            db.session.commit()
                            print(f"   💾 Salvando no banco...")
                            
                    except Exception as e:
                        erros += 1
                        if erros <= 5:  # Mostra apenas os primeiros 5 erros
                            print(f"   ⚠️  Erro na linha {linha_num}: {str(e)}")
                        continue
                
                # Commit final
                db.session.commit()
                
                print(f"\n✅ IMPORTAÇÃO CONCLUÍDA!")
                print(f"📊 Ramais importados: {ramais_importados}")
                print(f"❌ Erros de processamento: {erros}")
                
                if ramais_importados == 0:
                    print(f"🚨 ALERTA: Nenhum ramal foi importado!")
                    print(f"💡 Verifique se as colunas do CSV estão corretas")
                
        except Exception as e:
            print(f"❌ Erro durante a importação: {str(e)}")
            import traceback
            traceback.print_exc()
            db.session.rollback()

def main():
    """Função principal"""
    print("=" * 60)
    print("🔄 SISTEMA DE IMPORTACAO CORRIGIDA - DMOB")
    print("=" * 60)
    print("⚠️  ATENÇÃO: Esta versão vai:")
    print("   - LIMPAR todos os ramais existentes")
    print("   - IMPORTAR todos os 906 ramais do CSV")
    print("   - Usar a coluna CORRETA 'Ramal/Estrada' para os nomes")
    print("=" * 60)
    
    confirmacao = input("❓ Continuar? (s/N): ")
    if confirmacao.lower() != 's':
        print("❌ Importação cancelada")
        return
    
    try:
        # Importa dados do CSV
        importar_dados_csv()
        
        print("\n🎉 Processo finalizado!")
        
    except Exception as e:
        print(f"\n💥 Erro crítico: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()