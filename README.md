# Exchange Cache

Command-line tool that downloads metadata for all your Exchange emails in JSON.

## Synopsis

```
dotnet run [-c|--config]
```

## Options

- `-c|--config <PATH>`

  Specify the configuration file to use (default: config.json).

## Configuration

- `email` (string) Exchange email account to connect to
- `username` (string) Username to log in with
- `password` (string) Password to log in with

## Example configuration

```json
{
  "email": "example@outlook.com",
  "username": "example@outlook.com",
  "password": ""
}
```
