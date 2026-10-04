using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllersWithViews();

var app = builder.Build();
app.UseDeveloperExceptionPage();
app.UseDefaultFiles();
app.UseStaticFiles();
app.MapDefaultControllerRoute();
app.Run();
