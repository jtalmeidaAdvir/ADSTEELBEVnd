using Primavera.Extensibility.Sales.Editors;
using Primavera.Extensibility.BusinessEntities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Primavera.Extensibility.BusinessEntities.ExtensibilityService.EventArgs;
using VndBE100;
using CmpBE100;


namespace ADSTEELBEVnd.Sales
{
    public class UiEditorVendas : EditorVendas
    {
        public string TipoDoc => this.DocumentoVenda.Tipodoc;
        public string Serie => this.DocumentoVenda.Serie;
        public int numDoc => this.DocumentoVenda.NumDoc;

        public override void TeclaPressionada(int KeyCode, int Shift, ExtensibilityEventArgs e)
        {

            // if tipo doc for igual a FA
            if (this.TipoDoc == "FA")
            {

                var entidade = this.DocumentoVenda.Entidade;

                if (entidade == "0001")
                {
                    // Perguntar ao utilizador
                    var resposta = PSO.Dialogos.MostraMensagem(
                        StdPlatBS100.StdBSTipos.TipoMsg.PRI_SimNao,
                        "Deseja criar automaticamente a VFA na empresa MetalCarib?",
                        StdPlatBS100.StdBSTipos.IconId.PRI_Questiona
                    );

                    if (resposta == StdPlatBS100.StdBSTipos.ResultMsg.PRI_Sim)
                    {
                        //PSO.MensagensDialogos.MostraMensagem(StdPlatBS100.StdBSTipos.TipoMsg.PRI_Detalhe, "A criar VFA na empresa MetalCarib...", StdPlatBS100.StdBSTipos.IconId.PRI_Informativo);
                        CriarVfaMetalCarib();
                    }
                }
            }
        }
        public override void AntesDeGravar(ref bool Cancel, ExtensibilityEventArgs e)
        {
            try
            {
                // if tipo doc for igual a FA
                if (this.TipoDoc == "FA") { 
                
                    var entidade = this.DocumentoVenda.Entidade;

                    if (entidade == "0001")
                    {
                        // Perguntar ao utilizador
                        var resposta = PSO.Dialogos.MostraMensagem(
                            StdPlatBS100.StdBSTipos.TipoMsg.PRI_SimNao,
                            "Deseja criar automaticamente a VFA na empresa MetalCarib?",
                            StdPlatBS100.StdBSTipos.IconId.PRI_Questiona
                        );

                        if (resposta == StdPlatBS100.StdBSTipos.ResultMsg.PRI_Sim)
                        {
                            CriarVfaMetalCarib();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Cancel = true;
                PSO.Dialogos.MostraMensagem(
                    StdPlatBS100.StdBSTipos.TipoMsg.PRI_Detalhe,
                    ex.Message,
                    StdPlatBS100.StdBSTipos.IconId.PRI_Critico
                );
            }

        }

        private void CriarVfaMetalCarib()
        {
            var pso = new StdPlatBS100.StdPlatBS();
            var bso = new ErpBS100.ErpBS();

            try
            {
                var conf = new StdPlatBS100.StdBSConfApl
                {
                    AbvtApl = "ERP",
                    Instancia = "default",
                    Utilizador = "Cegid",
                    PwdUtilizador = "AdvirPlan@",
                    LicVersaoMinima = "10.00"
                };

                var trans = new StdBE100.StdBETransaccao();

                pso.AbrePlataformaEmpresa(
                    "METALCARIB",
                    trans,
                    conf,
                    StdBE100.StdBETipos.EnumTipoPlataforma.tpEvolution
                );

                bso.AbreEmpresaTrabalho(
                    StdBE100.StdBETipos.EnumTipoPlataforma.tpEvolution,
                    "METALCARIB",
                    conf.Utilizador,
                    conf.PwdUtilizador,
                    trans,
                    conf.Instancia
                );

                var numDocExterno = $"{TipoDoc}.{Serie}/{numDoc}";

                // VERIFICA SE JÁ EXISTE
                string sqlVerifica =
                    $"SELECT COUNT(*) AS Total " +
                    $"FROM [PRIMETALCARIB].[dbo].CabecCompras " +
                    $"WHERE NumDocExterno = '{numDocExterno}' AND Tipodoc = 'VFA'";

                var resultado = bso.Consulta(sqlVerifica);

                if (resultado != null && resultado.DaValor<int>("Total") > 0)
                {
                    PSO.Dialogos.MostraMensagem(
                        StdPlatBS100.StdBSTipos.TipoMsg.PRI_Detalhe,
                        $"Já existe uma VFA com o NumDocExterno {numDocExterno}.",
                        StdPlatBS100.StdBSTipos.IconId.PRI_Critico
                    );
                    return;
                }

                CmpBEDocumentoCompra doc = new CmpBEDocumentoCompra
                {
                    Tipodoc = "VFA",
                    DataDoc = DateTime.Today,
                    TipoEntidade = "F",
                    Entidade = "0001",
                    Serie = "2025",
                    NumDocExterno = numDocExterno
                };

                bso.Compras.Documentos.PreencheDadosRelacionados(doc);

                // LINHAS
                for (int i = 1; i <= this.DocumentoVenda.Linhas.NumItens; i++)
                {
                    var linhaVenda = this.DocumentoVenda.Linhas.GetEdita(i);
                    if (string.IsNullOrWhiteSpace(linhaVenda.Artigo))
                        continue; // ignora linhas sem artigo
                    double quantidade = linhaVenda.Quantidade;
                    string armazem = "A001";
                    string localizacao = "";
                    double precoUnitario = linhaVenda.PrecUnit;
                    double desconto1 = linhaVenda.Desconto1;
                    string lote = linhaVenda.Lote;
                    double descEntidade = this.DocumentoVenda.DescEntidade;
                    double precoTaxaIva = linhaVenda.TaxaIva;

                    int linhasAntes = doc.Linhas.NumItens;

                    try
                    {
                        bso.Compras.Documentos.AdicionaLinha(
                            doc,
                            linhaVenda.Artigo,
                            ref quantidade,
                            ref armazem,
                            ref localizacao,
                            precoUnitario,
                            desconto1,
                            lote,
                            linhaVenda.VariavelA,
                            linhaVenda.VariavelB,
                            linhaVenda.VariavelB,
                            descEntidade,
                            0,
                            this.DocumentoVenda.Arredondamento,
                            this.DocumentoVenda.ArredondamentoIva,
                            false,
                            false,
                            ref precoTaxaIva
                        );
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(
                            $"Erro ao adicionar o artigo '{linhaVenda.Artigo}' (linha {i}).\n{ex.Message}"
                        );
                    }

                    // validação crítica
                    if (doc.Linhas.NumItens <= linhasAntes)
                    {
                        throw new Exception(
                            $"O artigo '{linhaVenda.Artigo}' não foi adicionado ao documento (linha {i})."
                        );
                    }


                    // ✅ validação crítica
                    if (doc.Linhas.NumItens <= linhasAntes)
                    {
                        throw new Exception($"Erro ao adicionar a linha {i} ({linhaVenda.Artigo}).");
                    }

                    var linhaDoc = doc.Linhas.GetEdita(doc.Linhas.NumItens);

                    linhaDoc.TaxaIva = linhaVenda.TaxaIva;
                    linhaDoc.CodIva = linhaVenda.CodIva;
                    linhaDoc.CodIvaEcotaxa = linhaVenda.CodIvaEcotaxa;
                    linhaDoc.Quantidade = linhaVenda.Quantidade;
                    linhaDoc.QuantFormula = linhaVenda.QuantFormula;
                   linhaDoc.Formula = linhaVenda.Formula; 
                }


                bso.Compras.Documentos.Actualiza(doc);

                // ✅ SUCESSO
                PSO.Dialogos.MostraMensagem(
                    StdPlatBS100.StdBSTipos.TipoMsg.PRI_Detalhe,
                    $"VFA criada com sucesso.\nDocumento externo: {numDocExterno}",
                    StdPlatBS100.StdBSTipos.IconId.PRI_Informativo
                );
            }
            catch (Exception ex)
            {
                // ❌ ERRO GLOBAL
                PSO.Dialogos.MostraMensagem(
                    StdPlatBS100.StdBSTipos.TipoMsg.PRI_Detalhe,
                    $"Erro ao criar a VFA:\n{ex.Message}",
                    StdPlatBS100.StdBSTipos.IconId.PRI_Critico
                );
            }
            finally
            {
                // 🔒 FECHA SEMPRE
                try
                {
                    bso.FechaEmpresaTrabalho();
                }
                catch { }
            }
        }


    }
}
