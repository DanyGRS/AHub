using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;


namespace AiderHubAtual.Models
{
    public class Context : DbContext
    {
        public Context(DbContextOptions<Context> options) : base(options)
        { 
            
        }
        public DbSet<Evento> Eventos { get; set; }
        public DbSet<Relatorio> Relatorios { get; set; }
    }
}
