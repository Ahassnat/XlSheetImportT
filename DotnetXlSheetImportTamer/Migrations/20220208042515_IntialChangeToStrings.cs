using Microsoft.EntityFrameworkCore.Migrations;

namespace DotnetXlSheetImportTamer.Migrations
{
    public partial class IntialChangeToStrings : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "PriceList",
                table: "NVPCiscos",
                nullable: false,
                oldClrType: typeof(double),
                oldType: "float");

            migrationBuilder.AlterColumn<string>(
                name: "MinDiscount",
                table: "NVPCiscos",
                nullable: false,
                oldClrType: typeof(double),
                oldType: "float");

            migrationBuilder.AlterColumn<string>(
                name: "DiscountPrice",
                table: "NVPCiscos",
                nullable: false,
                oldClrType: typeof(double),
                oldType: "float");

            migrationBuilder.AlterColumn<string>(
                name: "Brand",
                table: "NVPCiscos",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<double>(
                name: "PriceList",
                table: "NVPCiscos",
                type: "float",
                nullable: false,
                oldClrType: typeof(string));

            migrationBuilder.AlterColumn<double>(
                name: "MinDiscount",
                table: "NVPCiscos",
                type: "float",
                nullable: false,
                oldClrType: typeof(string));

            migrationBuilder.AlterColumn<double>(
                name: "DiscountPrice",
                table: "NVPCiscos",
                type: "float",
                nullable: false,
                oldClrType: typeof(string));

            migrationBuilder.AlterColumn<int>(
                name: "Brand",
                table: "NVPCiscos",
                type: "int",
                nullable: false,
                oldClrType: typeof(string));
        }
    }
}
