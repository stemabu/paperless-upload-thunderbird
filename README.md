# Send to Paperless-ngx

<p align="center">
  <img src="icons/icon-512.png" alt="Send-to-Paperless-Icon" width="256" height="256">
</p>
<p align="center">This Addon is based on https://github.com/sebastian-xyz/paperless-upload-thunderbird, **but it was heavily customized to my personal needs**.</p>

<p align="center">**Ab Version 1.0.0 funktioniert das Addon nur noch mit Paperless-ngx V3!**</p>

---

## Benutzung

1. Öffne eine E-Mail (egal, ob mit oder ohne Attachment).
2. Klicke auf den **Send to Paperless-ngx**-Button in the Thunderbird-Toolbar.
3. Konfiguriere das Addon, indem Du die Paperless-ngx-Server-URL und den API-Key in den Optionen angibst (nur beim ersten Start notwendig).
4. Schließe die Optionen.
5. Durch Rechtsklick auf eine E-Mail in der E-Mail-Liste oder Drücken des Addon-Buttons erscheint der Upload-Dialog. Hier können der Korrespondent, ein Tag, die Richtung (eingehende oder ausgehende Mail), der Timeout beim Hochladen sowie ggf. mit zu übertragende Anhänge ausgewählt werden.
6. Nachdem eine E-Mail erfolgreich zu Paperless übertragen wurde, wird die E-Mail in Thunderbird mit dem Schlagwort "Paperless" versehen.


## Überblick

**Send to Paperless-ngx** ist ein Thunderbird-Add-on, das das Hochladen von E-Mails und deren Anhängen direkt auf Ihren [paperless-ngx](https://github.com/paperless-ngx/paperless-ngx)-Server vereinfacht. Mit wenigen Klicks können Sie Dokumente aus Ihrem Posteingang an Ihr Dokumentenmanagementsystem senden – ganz ohne manuelles Hochladen über die Paperless-ngx-Oberfläche.

---

## Features und Besonderheiten

- Hochladen von E-Mails und/oder Anhängen direkt in Paperless-ngx.
- Sichere, lokale Verarbeitung – keine Server von Drittanbietern.
- Eine hochgeladene E-Mail wird automatisch innerhalb von Paperless über die Funktion "dazugehörige Dokumente" mit den hochgeladenen Anhängen verknüpft und vice versa. Zusätzlich erfolgt eine Verknüpfung der Anhänge untereinander.
- Optional können in den Optionen Beziehungen zwischen E-Mail-Adressen und Korrespondenten erstellt werden, die dann im Upload-Dialog vorausgewählt werden.
- Es ist vorgesehen, für alle E-Mails eine Richtung anzugeben ("Eingang" oder "Ausgang").
- Möglichkeit, einen Time-Out festzulegen, falls der Server langsam arbeitet (default: 2 Minuten; 3, 4, 5, 10 Minuten).
- Möglichkeit, die Umwandlung der E-Mail zum PDF lokal vorzunehmen und erst dann zu Paperless zu senden (kann z.B. sinnvoll sein, wenn Paperless-ngx auf einem NAS läuft).
- Automatische Schlagwortvergabe "Paperless" in Thunderbird nach erfolgreichem Upload.
- [in Vorbereitung:] Upload mehrerer E-Mails gleichzeitig (nur für E-Mails ohne Anhänge)
- Deutsche Bedienoberfläche.

---

## Installation
### Default
- Create an XPI and install it in Thunderbird.

---

## Usage

1. Open an email with a PDF attachment.
2. Click the **Send to Paperless** icon in the Thunderbird toolbar.
3. Configure your paperless-ngx server URL and API key in the add-on’s options (first use only).
4. Rigth click the message and select the upload option.
5. Receive a notification when the upload is complete.

---

## Configuration

Go to **Add-ons and Themes** > **Extensions** > **Send to Paperless** > **Preferences** to set:

- **Server URL**: The base URL of your paperless-ngx instance
- **API Key**: Your personal API key for authentication

---

## Development

1. Clone this repository:
   ```bash
   git clone https://github.com/stemabu/send-to-paperless
   ```
2. Open the folder in VS Code or your preferred editor.
3. Make your changes and test the add-on in Thunderbird’s debug mode.

---


## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
