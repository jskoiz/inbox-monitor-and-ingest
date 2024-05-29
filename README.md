# Spreadsheet Email Handler to MongoDB

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Utility to monitor an inbox for new Excel files, automatically convert them to CSV format, transfer them into a MongoDB database, and serve them to end users through an easy-to-read web interface.

## Features

- Automatic monitoring of the inbox for new Excel files
- Conversion of Excel files to CSV format
- Storage of converted CSV files in a MongoDB database
- Web interface for end users to access and visualize the converted data

## Setup
Before running the application, make sure to set up the following environment variables:

- `EMAIL_USER`: Your email address or username for the monitored inbox.
- `EMAIL_PASSWORD`: Your password for the monitored inbox.
- `EMAIL_HOST`: The email server hostname or IP address.
- `EMAIL_PORT`: The email server port.
- `EMAIL_TLS`: Set to `true` if TLS/SSL should be used; otherwise, set to `false`.
- `MONGODB_URI`: The connection URI for your MongoDB database.
- `MONGODB_DB`: The name of the database to use.
- `MONGODB_COLLECTION`: The name of the collection to use.
- `MONGODB_USERNAME`: The username for your MongoDB database.
- `MONGODB_PASSWORD`: The password for your MongoDB database.
- `MONGODB_AUTH_SOURCE`: The authentication source for your MongoDB database.
- `MONGODB_AUTH_MECHANISM`: The authentication mechanism for your MongoDB database.
- `MONGODB_SSL`: Set to `true` if SSL should be used; otherwise, set to `false`.

## License

This project is licensed under the [MIT License](LICENSE).

## Contact

For questions, suggestions, or feedback, please contact [me](mailto:jerry@avmillabs.com).
