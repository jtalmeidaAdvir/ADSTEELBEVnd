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

            var conf = new StdPlatBS100.StdBSConfApl
            {
                AbvtApl = "ERP",
                Instancia = "default",
                Utilizador = "Cegid",
                PwdUtilizador = "AdvirPlan@",
                LicVersaoMinima = "10.00"
            };

            var trans = new StdBE100.StdBETransaccao();

            pso.AbrePlataformaEmpresa("CLONEMC2", trans, conf,
                StdBE100.StdBETipos.EnumTipoPlataforma.tpEvolution);

            bso.AbreEmpresaTrabalho(
                StdBE100.StdBETipos.EnumTipoPlataforma.tpEvolution,
                "CLONEMC2",
                conf.Utilizador,
                conf.PwdUtilizador,
                trans,
                conf.Instancia
            );

            var numDocExterno = $"{TipoDoc}.{Serie}/{numDoc}";
            CmpBEDocumentoCompra doc = new CmpBEDocumentoCompra();
            doc.Tipodoc = "VFA";
            doc.DataDoc = DateTime.Today;
            doc.TipoEntidade = "F";
            doc.Entidade = "0001";
            doc.NumDocExterno = numDocExterno;
            bso.Compras.Documentos.PreencheDadosRelacionados(doc);

            // Iterar sobre as linhas
            for (int i = 1; i <= this.DocumentoVenda.Linhas.NumItens; i++)
            {
                var linhaVenda = this.DocumentoVenda.Linhas.GetEdita(i);

                double quantidade = linhaVenda.Quantidade;
                string armazem = "A001";
                string localizacao = "";
                double precoUnitario = linhaVenda.PrecUnit;
                double desconto1 = linhaVenda.Desconto1;
                string lote = linhaVenda.Lote;
                double descEntidade = this.DocumentoVenda.DescEntidade;
                double precoTaxaIva = linhaVenda.TaxaIva;
                int arredondamento = this.DocumentoVenda.Arredondamento;
                int arredondaIva = this.DocumentoVenda.ArredondamentoIva;
                string obraid = linhaVenda.IDObra;
                double varA = linhaVenda.VariavelA;
                double varB = linhaVenda.VariavelB;
                double varC = linhaVenda.VariavelC;
                double quantFormula = linhaVenda.QuantFormula;
                

                bso.Compras.Documentos.AdicionaLinha(
                    doc,
                    linhaVenda.Artigo,
                    ref quantidade,
                    ref armazem,
                    ref localizacao,
                    precoUnitario,
                    desconto1,
                    lote,
                    varA, varB, varC,
                    descEntidade,
                    0,
                    arredondamento,
                    arredondaIva,
                    false,
                    false,
                    ref precoTaxaIva
                );
       

                // CORREÇÃO: Verificar se a linha foi realmente adicionada
                if (doc.Linhas.NumItens > 0)
                {
                    try
                    {
                        // Obter o índice correto da última linha
                        int indiceUltimaLinha = doc.Linhas.NumItens;
                        var linhaDoc = doc.Linhas.GetEdita(indiceUltimaLinha);

                        //se existir artigo
                  

                            if (!string.IsNullOrEmpty(linhaVenda.Artigo))
                            {
                            linhaDoc.TaxaIva = linhaVenda.TaxaIva;
                            linhaDoc.CodIva = linhaVenda.CodIva;
                            linhaDoc.CodIvaEcotaxa = linhaVenda.CodIvaEcotaxa;
                            linhaDoc.Quantidade = linhaVenda.Quantidade;
                        }

     

                       
                        // Buscar obra apenas se IDObra não estiver vazia
                        if (!string.IsNullOrEmpty(obraid))
                        {
                            var query1 = $"SELECT Codigo FROM [PRISTEELBE].[dbo].COP_Obras WHERE ID = '{obraid}'";
                            var resultadoObra = bso.Consulta(query1);

                            if (resultadoObra != null && resultadoObra.Vazia() == false)
                            {
                                var codigoObra = resultadoObra.DaValor<string>("Codigo");

                                var querybuscarCodigoobra = $"SELECT ID FROM [PRICLONEMC2].[dbo].[COP_Obras] WHERE Codigo = '{codigoObra}'";
                                var resultadoId = bso.Consulta(querybuscarCodigoobra);

                                if (resultadoId != null && resultadoId.Vazia() == false)
                                {
                                    var id = resultadoId.DaValor<string>("ID");
                                    linhaDoc.IDObra = id;
                                }
                            }
                        }

                        linhaDoc.QuantFormula = quantFormula;
                    }
                    catch (Exception ex)
                    {
                        // Log do erro específico da linha
                        PSO.Dialogos.MostraMensagem(
                            StdPlatBS100.StdBSTipos.TipoMsg.PRI_Detalhe,
                            $"Erro ao processar linha {i}: {ex.Message}",
                            StdPlatBS100.StdBSTipos.IconId.PRI_Critico
                        );
                    }
                }
            }

            bso.Compras.Documentos.Actualiza(doc);
            bso.FechaEmpresaTrabalho();
        }

    }
}
