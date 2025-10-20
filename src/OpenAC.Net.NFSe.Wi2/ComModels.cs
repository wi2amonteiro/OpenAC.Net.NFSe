using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OpenAC.Net.NFSe.Nota;

namespace OpenAC.Net.NFSe.Wi2
{
    [ComVisible(true)]
    [Guid("11111111-1111-1111-1111-111111111111")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComItem
    {
        public string Descricao { get; set; }
        public decimal Quantidade { get; set; }
        public decimal ValorUnitario { get; set; }
        public decimal ValorTotal { get; set; }
        public bool Tributavel { get; set; }
        public string Unidade { get; set; }
    }

    [ComVisible(true)]
    [Guid("22222222-2222-2222-2222-222222222222")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComValores
    {
        public decimal ValorServicos { get; set; }
        public decimal ValorDeducoes { get; set; }
        public decimal ValorPis { get; set; }
        public decimal ValorCofins { get; set; }
        public decimal ValorInss { get; set; }
        public decimal ValorIr { get; set; }
        public decimal ValorCsll { get; set; }
        public decimal ValorOutrasRetencoes { get; set; }
        public decimal BaseCalculo { get; set; }
        public decimal Aliquota { get; set; }
        public decimal ValorIss { get; set; }
        public decimal ValorLiquidoNfse { get; set; }
        public decimal ValorIssRetido { get; set; }
        public decimal DescontoCondicionado { get; set; }
        public decimal DescontoIncondicionado { get; set; }
        public bool IssRetido { get; set; }
    }

    [ComVisible(true)]
    [Guid("33333333-3333-3333-3333-333333333333")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComServico
    {
        public string ItemListaServico { get; set; }
        public string CodigoTributacaoMunicipio { get; set; }
        public string CodigoCnae { get; set; }
        public string Discriminacao { get; set; }
        public int CodigoMunicipio { get; set; }
        public ComValores Valores { get; set; } = new ComValores();

        private readonly List<ComItem> _itens = new List<ComItem>();
        public ComItem AddItem()
        {
            var it = new ComItem();
            _itens.Add(it);
            return it;
        }
        public ComItem[] GetItens() => _itens.ToArray();
    }

    [ComVisible(true)]
    [Guid("44444444-4444-4444-4444-444444444444")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComTomador
    {
        public string CpfCnpj { get; set; }
        public string InscricaoMunicipal { get; set; }
        public string RazaoSocial { get; set; }
        public string Logradouro { get; set; }
        public string Numero { get; set; }
        public string Complemento { get; set; }
        public string Bairro { get; set; }
        public int CodigoMunicipio { get; set; }
        public string Municipio { get; set; }
        public string Uf { get; set; }
        public string Cep { get; set; }
        public string Email { get; set; }
        public string DDD { get; set; }
        public string Telefone { get; set; }
    }

    [ComVisible(true)]
    [Guid("55555555-5555-5555-5555-555555555555")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComPrestador
    {
        public string CpfCnpj { get; set; }
        public string InscricaoMunicipal { get; set; }
        public string RazaoSocial { get; set; }
    }

    [ComVisible(true)]
    [Guid("66666666-6666-6666-6666-666666666666")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComIdentificacaoRps
    {
        public string Numero { get; set; }
        public string Serie { get; set; }
        public int Tipo { get; set; }
        public DateTime DataEmissao { get; set; } = DateTime.Now;
    }

    [ComVisible(true)]
    [Guid("77777777-7777-7777-7777-777777777777")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComNota
    {
        public ComIdentificacaoRps IdentificacaoRps { get; set; } = new ComIdentificacaoRps();
        public ComPrestador Prestador { get; set; } = new ComPrestador();
        public ComTomador Tomador { get; set; } = new ComTomador();
        public ComServico Servico { get; set; } = new ComServico();
        public DateTime Competencia { get; set; } = DateTime.Now;
        public string Situacao { get; set; } = "Normal";
        public bool OptanteSimplesNacional { get; set; } = true;
        public bool OptanteMEISimei { get; set; } = false;
        public string NaturezaOperacao { get; set; }
        public string RegimeEspecialTributacao { get; set; }
        public bool IncentivadorCultural { get; set; } = false;
        public string OutrasInformacoes { get; set; }
        public decimal ValorCredito { get; set; }

        private readonly List<ComItem> _itens = new List<ComItem>();
        public ComItem AddItem()
        {
            var it = new ComItem();
            _itens.Add(it);
            Servico?.GetType(); // prevent warning
            return it;
        }
        public ComItem[] GetItens() => _itens.ToArray();
    }
}
