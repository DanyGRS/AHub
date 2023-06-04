using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AiderHubAtual.Models
{
    [Table("relatorio")]
    public class Relatorio
    {
        [Key]
        [Column("id_relatorio")]
        public int IdRelatorio { get; set; }
        [Column("foto_logo")]
        public byte[] FotoLogo { get; set; }
        [Column("assinatura_digital")]
        public byte[] AssinaturaDigital { get; set; }
        [Column("nome_voluntario")]
        public string NomeVoluntario { get; set; }
        [Column("data_acao")]
        public DateTime DataAcao { get; set; }
        [Column("carga_horaria")]
        public TimeSpan CargaHoraria { get; set; }
        [Column("data_geracao")]
        public DateTime DataGeracao { get; set; }

        public Relatorio() { }

        public Relatorio(int idRelatorio, byte[] fotoLogo, byte[] assinaturaDigital, string nomeVoluntario, DateTime dataAcao, TimeSpan cargaHoraria, DateTime dataGeracao)
        {
            IdRelatorio = idRelatorio;
            FotoLogo = fotoLogo;
            AssinaturaDigital = assinaturaDigital;
            NomeVoluntario = nomeVoluntario;
            DataAcao = dataAcao;
            CargaHoraria = cargaHoraria;
            DataGeracao = dataGeracao;
        }
    }
}
