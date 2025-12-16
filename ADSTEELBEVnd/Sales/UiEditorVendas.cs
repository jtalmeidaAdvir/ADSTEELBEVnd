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

            // Abrir plataforma e empresa MetalCarib
            pso.AbrePlataformaEmpresa("METALCARIB", trans, conf,
                StdBE100.StdBETipos.EnumTipoPlataforma.tpEvolution);

            bso.AbreEmpresaTrabalho(
                StdBE100.StdBETipos.EnumTipoPlataforma.tpEvolution,
                "METALCARIB",
                conf.Utilizador,
                conf.PwdUtilizador,
                trans,
                conf.Instancia
            );


            var numDocExterno = $"{TipoDoc}.{Serie}/{numDoc}";
            // Criar documento
            CmpBEDocumentoCompra doc = new CmpBEDocumentoCompra();
            doc.Tipodoc = "VFA";
            doc.DataDoc = DateTime.Today;
            doc.TipoEntidade = "F"; // fornecedor
            doc.Entidade = "0001"; // fornecedor
            doc.NumDocExterno = numDocExterno;
            bso.Compras.Documentos.PreencheDadosRelacionados(doc);


            // Iterar sobre as linhas da venda e copiar para a VFA
            // Iterar sobre as linhas da venda
            for (int i = 1; i <= this.DocumentoVenda.Linhas.NumItens; i++)
            {
                var linhaVenda = this.DocumentoVenda.Linhas.GetEdita(i);

                double quantidade = linhaVenda.Quantidade;
                string armazem = linhaVenda.Armazem; // ou mapear a partir da linha da venda
                string localizacao = linhaVenda.Localizacao; // opcional
                double precoUnitario = linhaVenda.PrecUnit;
                double desconto1 = linhaVenda.Desconto1;
                string lote = linhaVenda.Lote;
                double descEntidade = this.DocumentoVenda.DescEntidade;
                double precoTaxaIva = linhaVenda.TaxaIva;
                int arredondamento = this.DocumentoVenda.Arredondamento;
                int arredondaIva = this.DocumentoVenda.ArredondamentoIva;

                // Adiciona a linha ao documento de compras
                bso.Compras.Documentos.AdicionaLinha(
                    doc,
                    linhaVenda.Artigo,
                    ref quantidade,
                    ref armazem,
                    ref localizacao,
                    precoUnitario,
                    desconto1,
                    lote,       // Lote
                    0, 0, 0, // QntVariavelA/B/C
                    descEntidade,       // DescEntidade
                    0,
                    arredondamento,       // Arredondamento
                    arredondaIva,       // ArredondaIva
                    false,   // AdicionaArtigoAssociado
                    false,   // PrecoIvaIncluido
                    ref precoTaxaIva
                );
            }

            // Gravar
            bso.Compras.Documentos.Actualiza(doc);

            // (Opcional) fechar empresa
            bso.FechaEmpresaTrabalho();
        }

    }
}
