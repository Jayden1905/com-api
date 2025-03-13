var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();
builder.Services.AddSingleton<com_api.Services.PowerPointService>(); // Register PowerPoint service

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseHttpsRedirection();

var summaries = new[]
{
    "Freezing",
    "Bracing",
    "Chilly",
    "Cool",
    "Mild",
    "Warm",
    "Balmy",
    "Hot",
    "Sweltering",
    "Scorching",
};

app.MapGet(
        "/weatherforecast",
        () =>
        {
            var forecast = Enumerable
                .Range(1, 5)
                .Select(index => new WeatherForecast(
                    DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                    Random.Shared.Next(-20, 55),
                    summaries[Random.Shared.Next(summaries.Length)]
                ))
                .ToArray();
            return forecast;
        }
    )
    .WithName("GetWeatherForecast");

// PowerPoint Controller Endpoints
app.MapGet(
        "/powerpoint/open",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (
            com_api.Services.PowerPointService pptService,
            string filePath,
            bool? startSlideShow = true,
            bool? readOnly = true
        ) =>
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return Results.BadRequest("File path is required");
            }

            bool success = pptService.OpenPresentation(
                filePath,
                startSlideShow ?? true,
                readOnly ?? true
            );
            if (success)
            {
                return Results.Ok(
                    new
                    {
                        Message = "Presentation opened successfully",
                        TotalSlides = pptService.GetTotalSlides(),
                    }
                );
            }

            return Results.BadRequest("Failed to open presentation");
        }
    )
    .WithName("OpenPresentation");

app.MapGet(
        "/powerpoint/next",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService) =>
        {
            // Check if a presentation is active first
            int totalSlides = pptService.GetTotalSlides();
            if (totalSlides <= 0)
            {
                return Results.BadRequest(
                    "No active presentation. Please open a presentation first."
                );
            }

            bool success = pptService.NextSlide();
            if (success)
            {
                return Results.Ok(
                    new
                    {
                        CurrentSlide = pptService.GetCurrentSlideNumber(),
                        TotalSlides = totalSlides,
                    }
                );
            }

            return Results.BadRequest("Failed to navigate to next slide");
        }
    )
    .WithName("NextSlide");

app.MapGet(
        "/powerpoint/previous",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService) =>
        {
            // Check if a presentation is active first
            int totalSlides = pptService.GetTotalSlides();
            if (totalSlides <= 0)
            {
                return Results.BadRequest(
                    "No active presentation. Please open a presentation first."
                );
            }

            bool success = pptService.PreviousSlide();
            if (success)
            {
                return Results.Ok(
                    new
                    {
                        CurrentSlide = pptService.GetCurrentSlideNumber(),
                        TotalSlides = totalSlides,
                    }
                );
            }

            return Results.BadRequest("Failed to navigate to previous slide");
        }
    )
    .WithName("PreviousSlide");

app.MapGet(
        "/powerpoint/goto/{slideNumber}",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService, int slideNumber) =>
        {
            // Check if a presentation is active first
            int totalSlides = pptService.GetTotalSlides();
            if (totalSlides <= 0)
            {
                return Results.BadRequest(
                    "No active presentation. Please open a presentation first."
                );
            }

            // Validate slide number
            if (slideNumber < 1 || slideNumber > totalSlides)
            {
                return Results.BadRequest(
                    $"Invalid slide number. Please specify a value between 1 and {totalSlides}."
                );
            }

            bool success = pptService.GoToSlide(slideNumber);
            if (success)
            {
                return Results.Ok(
                    new
                    {
                        CurrentSlide = pptService.GetCurrentSlideNumber(),
                        TotalSlides = totalSlides,
                    }
                );
            }

            return Results.BadRequest($"Failed to navigate to slide {slideNumber}");
        }
    )
    .WithName("GoToSlide");

app.MapGet(
        "/powerpoint/status",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService) =>
        {
            int currentSlide = pptService.GetCurrentSlideNumber();
            int totalSlides = pptService.GetTotalSlides();

            if (currentSlide >= 0 && totalSlides > 0)
            {
                return Results.Ok(new { CurrentSlide = currentSlide, TotalSlides = totalSlides });
            }

            return Results.BadRequest("No active presentation");
        }
    )
    .WithName("PresentationStatus");

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

// Add a diagnostic endpoint
app.MapGet(
        "/powerpoint/diagnostics",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService, string? filePath = null) =>
        {
            string status = pptService.CheckPowerPointStatus(filePath);
            return Results.Ok(new { Diagnostics = status });
        }
    )
    .WithName("PowerPointDiagnostics");

// Add a force-quit endpoint to ensure PowerPoint closes properly
app.MapGet(
        "/powerpoint/force-quit",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService) =>
        {
            bool success = pptService.ForceQuitPowerPoint();
            if (success)
            {
                return Results.Ok(new { Message = "PowerPoint forcefully closed" });
            }
            return Results.BadRequest("Failed to force-quit PowerPoint");
        }
    )
    .WithName("ForceQuitPowerPoint");

// Add a preload endpoint for faster subsequent opening
app.MapGet(
        "/powerpoint/preload",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService, string filePath) =>
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return Results.BadRequest("File path is required");
            }

            bool success = pptService.PreloadPresentation(filePath);
            if (success)
            {
                return Results.Ok(new { Message = "Presentation preloaded" });
            }
            return Results.BadRequest("Failed to preload presentation");
        }
    )
    .WithName("PreloadPresentation");

// Add a quick reopen endpoint to rapidly reopen the last presentation
app.MapGet(
        "/powerpoint/reopen",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        (com_api.Services.PowerPointService pptService, bool? startSlideShow = true) =>
        {
            bool success = pptService.ReopenLastPresentation(startSlideShow ?? true);
            if (success)
            {
                return Results.Ok(
                    new
                    {
                        Message = "Presentation reopened successfully",
                        TotalSlides = pptService.GetTotalSlides(),
                    }
                );
            }
            return Results.BadRequest("Failed to reopen last presentation");
        }
    )
    .WithName("ReopenPresentation");

// Add a reset endpoint to force restart PowerPoint
app.MapGet(
        "/powerpoint/reset",
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        async (com_api.Services.PowerPointService pptService) =>
        {
            // First force quit PowerPoint
            pptService.ForceQuitPowerPoint();

            // Wait a moment for resources to be released
            await Task.Delay(1000);

            // Get diagnostics after reset
            string status = pptService.CheckPowerPointStatus();

            return Results.Ok(new { Message = "PowerPoint reset attempted", Status = status });
        }
    )
    .WithName("ResetPowerPoint");

app.Run();

record WeatherForecast(DateOnly Date, int TemperatureC, string? Summary)
{
    public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);
}
