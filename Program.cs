using System.Net;
using Microsoft.Extensions.Options;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddSingleton<com_api.Services.PowerPointService>();

builder.Services.AddCors(options =>
{
    options.AddPolicy(
        "AllowAll",
        policy =>
        {
            policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader();
        }
    );
});

builder.WebHost.ConfigureKestrel(options =>
{
    options.Listen(
        IPAddress.Any,
        5001,
        listenOptions =>
        {
            listenOptions.UseHttps("vm-api.crt", "vm-api.key");
        }
    );
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("AllowAll");
app.UseHttpsRedirection();

// PowerPoint Controller Endpoints
app.MapGet(
        "/powerpoint/open",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (
            com_api.Services.PowerPointService pptService,
            string filePath,
            bool? startSlideShow = true
        ) =>
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return Results.BadRequest("File path is required");
            }

            bool success = pptService.OpenPresentation(filePath, startSlideShow ?? true);
            if (success)
            {
                return Results.Ok(new { Message = "Presentation opened successfully" });
            }

            return Results.BadRequest("Failed to open presentation");
        }
    )
    .WithName("OpenPresentation");

app.MapGet(
        "/powerpoint/goto/{slideNumber}",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService, int slideNumber) =>
        {
            bool success = pptService.GoToSlide(slideNumber);
            if (success)
            {
                return Results.Ok(new { Message = $"Navigated to slide {slideNumber}" });
            }

            return Results.BadRequest($"Failed to navigate to slide {slideNumber}");
        }
    )
    .WithName("GoToSlide");

app.MapGet(
        "/powerpoint/close",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService) =>
        {
            bool success = pptService.ClosePresentation();
            if (success)
            {
                return Results.Ok(new { Message = "Presentation closed successfully" });
            }

            return Results.BadRequest("Failed to close presentation or no active presentation");
        }
    )
    .WithName("ClosePresentation");

app.Run();
