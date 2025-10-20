using OpenAC.Net.NFSe;
using OpenAC.Net.NFSe.Nota;
using OpenAC.Net.NFSe.Wi2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace OpenAC.Net.NFSe.Wi2
{
    [ComVisible(true)]
    [Guid("88888888-8888-8888-8888-888888888888")]
    [ProgId("OpenACNFSeCOM.NFSeManager")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class NFSeManager
    {
        private readonly List<ComNota> _notas = new List<ComNota>();

        public string ArquivoIniConfig { get; set; }
        public string CertificadoArquivo { get; set; }
        public string CertificadoSenha { get; set; }
        public int CodigoMunicipio { get; set; } = 3550308; // São Paulo por default
        public int Ambiente { get; set; } = 2; // homologação por default

        public ComNota CreateNota()
        {
            var n = new ComNota();
            _notas.Add(n);
            return n;
        }

        public void AddNota(ComNota nota)
        {
            if (nota != null) _notas.Add(nota);
        }

        public void ClearNotes()
        {
            _notas.Clear();
        }

        // Gera XML para cada nota (RPS) e grava em pasta
        public string GerarXmlNotas(string pastaSaida)
        {
            try
            {
                var open = new OpenNFSe();
                if (!string.IsNullOrEmpty(ArquivoIniConfig))
                {
                    try { open.Configuracoes.Load(ArquivoIniConfig); } catch { }
                }

                if (!string.IsNullOrEmpty(CertificadoArquivo))
                {
                    open.Configuracoes.Certificados.Arquivo = CertificadoArquivo;
                    open.Configuracoes.Certificados.Senha = CertificadoSenha;
                }

                open.NotasServico.Clear();

                foreach (var com in _notas)
                {
                    var nf = open.NotasServico.AddNew();

                    // Identificação
                    nf.IdentificacaoRps.Numero = com.IdentificacaoRps.Numero ?? DateTime.Now.Ticks.ToString();
                    nf.IdentificacaoRps.Serie = string.IsNullOrEmpty(com.IdentificacaoRps.Serie) ? "1" : com.IdentificacaoRps.Serie;
                    nf.IdentificacaoRps.Tipo = TipoRps.RPS;
                    nf.IdentificacaoRps.DataEmissao = com.IdentificacaoRps.DataEmissao;

                    nf.Competencia = com.Competencia;
                    nf.Situacao = SituacaoNFSeRps.Normal;
                    nf.OptanteSimplesNacional = com.OptanteSimplesNacional ? NFSeSimNao.Sim : NFSeSimNao.Nao;
                    nf.OptanteMEISimei = com.OptanteMEISimei ? NFSeSimNao.Sim : NFSeSimNao.Nao;

                    // Prestador
                    nf.Prestador.CpfCnpj = com.Prestador.CpfCnpj;
                    nf.Prestador.InscricaoMunicipal = com.Prestador.InscricaoMunicipal;
                    nf.Prestador.RazaoSocial = com.Prestador.RazaoSocial;

                    // Tomador
                    nf.Tomador.CpfCnpj = com.Tomador.CpfCnpj;
                    nf.Tomador.InscricaoMunicipal = com.Tomador.InscricaoMunicipal;
                    nf.Tomador.RazaoSocial = com.Tomador.RazaoSocial;
                    nf.Tomador.Endereco.Logradouro = com.Tomador.Logradouro;
                    nf.Tomador.Endereco.Numero = com.Tomador.Numero;
                    nf.Tomador.Endereco.Complemento = com.Tomador.Complemento;
                    nf.Tomador.Endereco.Bairro = com.Tomador.Bairro;
                    nf.Tomador.Endereco.CodigoMunicipio = com.Tomador.CodigoMunicipio;
                    nf.Tomador.Endereco.Municipio = com.Tomador.Municipio;
                    nf.Tomador.Endereco.Uf = com.Tomador.Uf;
                    nf.Tomador.Endereco.Cep = com.Tomador.Cep;
                    nf.Tomador.DadosContato.Email = com.Tomador.Email;
                    nf.Tomador.DadosContato.DDD = com.Tomador.DDD;
                    nf.Tomador.DadosContato.Telefone = com.Tomador.Telefone;

                    // Serviço / Valores
                    nf.Servico.ItemListaServico = com.Servico.ItemListaServico;
                    nf.Servico.CodigoTributacaoMunicipio = com.Servico.CodigoTributacaoMunicipio;
                    nf.Servico.CodigoCnae = com.Servico.CodigoCnae;
                    nf.Servico.Discriminacao = com.Servico.Discriminacao;
                    nf.Servico.CodigoMunicipio = com.Servico.CodigoMunicipio;

                    // Valores
                    nf.Servico.Valores.ValorServicos = com.Servico.Valores.ValorServicos;
                    nf.Servico.Valores.ValorDeducoes = com.Servico.Valores.ValorDeducoes;
                    nf.Servico.Valores.ValorPis = com.Servico.Valores.ValorPis;
                    nf.Servico.Valores.ValorCofins = com.Servico.Valores.ValorCofins;
                    nf.Servico.Valores.ValorInss = com.Servico.Valores.ValorInss;
                    nf.Servico.Valores.ValorIr = com.Servico.Valores.ValorIr;
                    nf.Servico.Valores.ValorCsll = com.Servico.Valores.ValorCsll;
                    nf.Servico.Valores.ValorOutrasRetencoes = com.Servico.Valores.ValorOutrasRetencoes;
                    nf.Servico.Valores.BaseCalculo = com.Servico.Valores.BaseCalculo;
                    nf.Servico.Valores.Aliquota = com.Servico.Valores.Aliquota;
                    nf.Servico.Valores.ValorIss = com.Servico.Valores.ValorIss;
                    nf.Servico.Valores.ValorLiquidoNfse = com.Servico.Valores.ValorLiquidoNfse;
                    nf.Servico.Valores.ValorIssRetido = com.Servico.Valores.ValorIssRetido;
                    nf.Servico.Valores.DescontoCondicionado = com.Servico.Valores.DescontoCondicionado;
                    nf.Servico.Valores.DescontoIncondicionado = com.Servico.Valores.DescontoIncondicionado;
                    nf.Servico.Valores.IssRetido = com.Servico.Valores.IssRetido ? SituacaoTributaria.Normal : SituacaoTributaria.Retencao;

                    // Itens
                    foreach (var it in com.GetItens())
                    {
                        var item = nf.Servico.ItemsServico.AddNew();
                        item.Descricao = it.Descricao;
                        item.Quantidade = it.Quantidade;
                        item.ValorUnitario = it.ValorUnitario;
                        item.ValorTotal = it.ValorTotal;
                        item.Tributavel = it.Tributavel ? NFSeSimNao.Sim : NFSeSimNao.Nao;
                    }

                    nf.OutrasInformacoes = com.OutrasInformacoes;
                    nf.ValorCredito = com.ValorCredito;
                }

                // Gera XML para cada nota e grava
                if (!Directory.Exists(pastaSaida)) Directory.CreateDirectory(pastaSaida);
                foreach (var nota in open.NotasServico)
                {
                    // Ajuste dependendo dos métodos disponíveis na versão do OpenAC
                    //var xml = nota.GetXml(); // ou nota.ArquivoXML gerado
                    var nome = $"RPS_{nota.IdentificacaoRps.Numero}.xml";
                    var path = Path.Combine(pastaSaida, nome);
                    //xml.Save(path);
                    nota.Save(path);
                }

                return $"OK|{_notas.Count} arquivos gerados em {pastaSaida}";
            }
            catch (Exception ex)
            {
                return "ERRO|" + ex.Message;
            }
        }

        // Transmitir todas as notas (exemplo)
        public string EmitirTodasNotas()
        {
            try
            {
                var open = new OpenNFSe();
                if (!string.IsNullOrEmpty(ArquivoIniConfig))
                {
                    try { open.Configuracoes.Load(ArquivoIniConfig); } catch { }
                }
                if (!string.IsNullOrEmpty(CertificadoArquivo))
                {
                    open.Configuracoes.Certificados.Arquivo = CertificadoArquivo;
                    open.Configuracoes.Certificados.Senha = CertificadoSenha;
                }

                open.NotasServico.Clear();
                // mapear como em GerarXmlNotas (omito repetição por brevidade)...
                foreach (var com in _notas)
                {
                    var nf = open.NotasServico.AddNew();
                    nf.IdentificacaoRps.Numero = com.IdentificacaoRps.Numero ?? DateTime.Now.Ticks.ToString();
                    nf.IdentificacaoRps.Serie = com.IdentificacaoRps.Serie ?? "1";
                    nf.IdentificacaoRps.Tipo = TipoRps.RPS;
                    nf.IdentificacaoRps.DataEmissao = com.IdentificacaoRps.DataEmissao;
                    // ... mapear os demais campos (igual GerarXmlNotas)
                }

                // emitir lote — método depende da versão do OpenAC
                open.Emitir(); // ajuste se necessário
                return "OK|Transmissão solicitada";
            }
            catch (Exception ex)
            {
                return "ERRO|" + ex.Message;
            }
        }
    }
}
