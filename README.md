PSDoc
=====

## What is PSDoc?

PSDoc is convention based PS tooling that wraps the Pandoc tool to provide a automateable document system.

Seperate your document into several smaller source files, write a config file ordering the files, and then generate a single document of any format that Pandoc supports

## Conventions

Convention based folder structure:

```
config
content
headers
output
```

Config file is a PowerShell hash structure that matches document order and placement

```
@{
    Name      = "Foo.1.3.Manual"
    "Header"  = 'headers/internal.md'
    "Content" = @{
        1 = 'content/overview.md'
        2 = 'content/server.md'
        3 = 'content/client.md'
        4 = 'content/validation.md'
        5 = 'content/architecture_long.md'
        6 = 'content/module_dev.md'
        7 = 'content/errata.md'
    }
    Formats = @(".docx",".html",".textile",".markdown")
}
```

## Execution

Will use the folder specified to generate the formatted document

` Format-Files -Path $home\code\foo-docs`