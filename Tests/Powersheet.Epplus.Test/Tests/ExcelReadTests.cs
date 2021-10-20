using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Nerosoft.Powersheet.Epplus.Test
{
    public class ExcelReadTests
    {
        private readonly ISheetWrapper _wrapper;

        public ExcelReadTests(ISheetWrapper wrapper)
        {
            _wrapper = wrapper;
        }

        [Fact]
        public async Task TestReadToDataTableAsync_CN()
        {
            var file = Path.Combine(AppContext.BaseDirectory, "Samples", "Employees.CN.xlsx");

            var options = new SheetReadOptions();

            options.AddMapProfile("Name", "姓名");
            options.AddMapProfile("Gender", "性别", (value, _) =>
            {
                return value switch
                {
                    null => 0,
                    "男" => 1,
                    "女" => 2,
                    _ => 0
                };
            });
            options.AddMapProfile("Age", "年龄");
            options.AddMapProfile("Birthdate", "出生日期");
            options.AddMapProfile("Department", "部门");
            options.AddMapProfile("IsActive", "是否在职", (value, _) => IsActiveValueConvert(value));

            var datatable = await _wrapper.ReadToDataTableAsync(file, options, 0);
            Assert.Equal(3, datatable.Rows.Count);
        }

        [Fact]
        public async Task TestReadToListAsync_CN()
        {
            var file = Path.Combine(AppContext.BaseDirectory, "Samples", "Employees.CN.xlsx");

            var options = new SheetReadOptions();

            options.AddMapProfile("Id", "编号");
            options.AddMapProfile("Name", "姓名");
            options.AddMapProfile("Gender", "性别", (value, _) =>
            {
                if (value is string gender)
                {
                    return gender switch
                    {
                        "男" => 1,
                        "女" => 2,
                        _ => 0
                    };
                }

                return value;
            });
            options.AddMapProfile("Age", "年龄");
            options.AddMapProfile("Birthdate", "出生日期");
            options.AddMapProfile("Department", "部门");
            options.AddMapProfile("IsActive", "是否在职", (value, _) => IsActiveValueConvert(value));

            var result = await _wrapper.ReadToListAsync<Employee>(file, options, 0);
            Assert.Equal(3, result.Count);
        }

        [Fact]
        public async Task TestReadToListAsync_CN_InvasiveMap()
        {
            var file = Path.Combine(AppContext.BaseDirectory, "Samples", "Employees.CN.xlsx");

            /*
            var options = new SheetReadOptions();

            options.AddMapProfile("Id", "编号");
            options.AddMapProfile("Gender", "性别", (value, _) =>
            {
                return value switch
                {
                    null => 0,
                    "男" => 1,
                    "女" => 2,
                    _ => 0
                };
            });
            options.AddMapProfile("Age", "年龄");
            options.AddMapProfile("Birthdate", "出生日期");
            options.AddMapProfile("IsActive", "是否在职", (value, _) => IsActiveValueConvert(value));
*/
            
            var builder = SheetHandleOptionsBuilder<Employee>.For<SheetReadOptions>()
                                                             .ConfigureProfile(item =>
                                                             {
                                                                 item.Property(t => t.Id).HasColumnName("编号");
                                                                 item.Property(t => t.Name).HasColumnName("姓名");
                                                                 item.Property(t => t.Gender).HasColumnName("性别")
                                                                     .HasValueConverter((value, _) =>
                                                                     {
                                                                         return value switch
                                                                         {
                                                                             null => 0,
                                                                             "男" => 1,
                                                                             "女" => 2,
                                                                             _ => 0
                                                                         };
                                                                     });
                                                                 item.Property(t => t.Birthdate).HasColumnName("出生日期");
                                                                 item.Property(t => t.Age).HasColumnName("年龄");
                                                                 item.Property(t => t.Department).HasColumnName("部门");
                                                                 item.Property(t => t.IsActive).HasColumnName("是否在职")
                                                                     .HasValueConverter((value, _) => IsActiveValueConvert(value));
                                                             });

            var options = builder.Options;
            var result = await _wrapper.ReadToListAsync<EmployeeInvasiveMap>(file, options, 0);
            Assert.Equal(3, result.Count);
        }
        
        private static bool IsActiveValueConvert(object value)
        {
            return value switch
            {
                "是" => true,
                "否" => false,
                "在职" => true,
                "离职" => false,
                _ => false
            };
        }
    }

    public class Employee
    {
        public Guid Id { get; set; }

        public string Name { get; set; }

        public int Gender { get; set; }

        public int Age { get; set; }

        public DateTime Birthdate { get; set; }

        public string Department { get; set; }

        public bool IsActive { get; set; }
    }
    
    public class EmployeeInvasiveMap
    {
        public Guid Id { get; set; }

        [SheetColumnMap("姓名")]
        public string Name { get; set; }

        public int Gender { get; set; }

        public int Age { get; set; }

        public DateTime Birthdate { get; set; }

        [SheetColumnMap("部门")]
        public string Department { get; set; }

        public bool IsActive { get; set; }
    }
}