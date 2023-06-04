using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AiderHubAtual.Models
{
    [Table("evento")]
    public class Evento
    {
        [Key]
        [Column("id_evento")]
        public int id_Evento { get; set; }
        //public System.DateTime data_Hora { get; set; }
        [Column("endereco")]
        public string Endereco { get; set; }
        // public System.TimeSpan Carga_Horario { get; set; }
        [Column("descricao")]
        public string Descricao { get; set; }
        [Column("responsavel")]
        public string Responsavel { get; set; }

        public Evento() { }

        public Evento(int id_evento, /*System.DateTime data_hora*/ string endereco,/*/System.TimeSpan carga_horario*/ string descricao, string responsavel)
        {
            id_Evento = id_evento;
            //data_Hora = data_hora;
            Endereco = endereco;
            //Carga_Horario = carga_horario;
            Descricao = descricao;
            Responsavel = responsavel;
        }
    }
}
