using HelloBlazor.Client.Pages;
using HelloBlazor.Components;
using System.Text;

var builder = WebApplication.CreateBuilder(args);

// Isso "desbloqueia" as codificações antigas do Windows para o RtfPipe funcionar
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddHubOptions(options =>
    {
        // Aumentando o limite para 64 MB (padrão é 32KB)
        options.MaximumReceiveMessageSize = 64 * 1024 * 1024;
    }).AddInteractiveWebAssemblyComponents();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseWebAssemblyDebugging();
}
else
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(HelloBlazor.Client._Imports).Assembly);

app.Run();
