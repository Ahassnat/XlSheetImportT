using Microsoft.EntityFrameworkCore.Migrations;

namespace DotnetXlSheetImportTamer.Migrations
{
    public partial class Intial : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "NVPCiscos",
                columns: table => new
                {
                    PartSKU = table.Column<string>(nullable: false),
                    CategoryCode = table.Column<string>(nullable: false),
                    Brand = table.Column<int>(nullable: false),
                    Manufacturer = table.Column<string>(nullable: false),
                    ItemDescription = table.Column<string>(nullable: false),
                    PriceList = table.Column<double>(nullable: false),
                    MinDiscount = table.Column<double>(nullable: false),
                    DiscountPrice = table.Column<double>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NVPCiscos", x => x.PartSKU);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "NVPCiscos");
        }
    }
}
