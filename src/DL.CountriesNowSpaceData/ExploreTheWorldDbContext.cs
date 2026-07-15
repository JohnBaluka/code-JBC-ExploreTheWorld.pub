using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData
{
    public class ExploreTheWorldDbContext : DbContext
    {
        public ExploreTheWorldDbContext(DbContextOptions<ExploreTheWorldDbContext> options) : base(options) { }

        // ================================================================
        // CNS Tables (CountriesNowSpaceApi)
        // ================================================================

        public DbSet<cns_Country> cns_Countries { get; set; }
        public DbSet<cns_City> cns_Cities { get; set; }
        public DbSet<cns_CountryCapital> cns_CountryCapitals { get; set; }
        public DbSet<cns_CountryFlag> cns_CountryFlags { get; set; }
        public DbSet<cns_CountryPopulation> cns_CountryPopulations { get; set; }
        public DbSet<cns_PopulationCount> cns_PopulationCounts { get; set; }
        public DbSet<cns_CountryStates> cns_CountryStates { get; set; }
        public DbSet<cns_CountryState> cns_CountryStateItems { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // ================================================================
            // CNS Tables (CountriesNowSpaceApi)
            // ================================================================

            modelBuilder.Entity<cns_Country>(entity =>
            {
                entity.ToTable("cns_Country");
                entity.HasKey(e => e.Iso2);
                entity.HasMany(e => e.Cities).WithOne(e => e.Country).HasForeignKey(e => e.Iso2);
                entity.HasOne(e => e.Capital).WithOne(e => e.Country).HasForeignKey<cns_CountryCapital>(e => e.Iso2);
                entity.HasOne(e => e.Flag).WithOne(e => e.Country).HasForeignKey<cns_CountryFlag>(e => e.Iso2);
            });

            modelBuilder.Entity<cns_City>(entity =>
            {
                entity.ToTable("cns_City");
                entity.HasKey(e => e.ID);
                entity.Property(e => e.ID).ValueGeneratedOnAdd();
                entity.HasIndex(e => e.Iso2);
            });

            modelBuilder.Entity<cns_CountryCapital>(entity =>
            {
                entity.ToTable("cns_CountryCapital");
                entity.HasKey(e => e.Iso2);
            });

            modelBuilder.Entity<cns_CountryFlag>(entity =>
            {
                entity.ToTable("cns_CountryFlag");
                entity.HasKey(e => e.Iso2);
            });

            modelBuilder.Entity<cns_CountryPopulation>(entity =>
            {
                entity.ToTable("cns_CountryPopulation");
                entity.HasKey(e => e.ID);
                entity.Property(e => e.ID).ValueGeneratedOnAdd();
                entity.HasMany(e => e.PopulationCounts).WithOne(e => e.CountryPopulation).HasForeignKey(e => e.CountryPopulation_ID);
            });

            modelBuilder.Entity<cns_PopulationCount>(entity =>
            {
                entity.ToTable("cns_PopulationCount");
                entity.HasKey(e => e.ID);
                entity.Property(e => e.ID).ValueGeneratedOnAdd();
                entity.HasIndex(e => e.CountryPopulation_ID);
            });

            modelBuilder.Entity<cns_CountryStates>(entity =>
            {
                entity.ToTable("cns_CountryStates");
                entity.HasKey(e => e.ID);
                entity.Property(e => e.ID).ValueGeneratedOnAdd();
                entity.HasMany(e => e.States).WithOne(e => e.CountryStates).HasForeignKey(e => e.CountryStates_ID);
            });

            modelBuilder.Entity<cns_CountryState>(entity =>
            {
                entity.ToTable("cns_CountryState");
                entity.HasKey(e => e.ID);
                entity.Property(e => e.ID).ValueGeneratedOnAdd();
                entity.HasIndex(e => e.CountryStates_ID);
            });
        }
    }
}
