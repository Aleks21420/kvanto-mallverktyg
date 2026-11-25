# Installera Kvanto Mallverktyg i Excel för Windows

Kopiera kommandot med knappen i rutans övre högra hörn och klistra in det i `cmd`.

1. Stäng Excel.
2. Tryck på **Windows-tangenten + R**, skriv `cmd` och tryck **Enter**.
3. Högerklicka i det svarta fönstret för att klistra in kommandot och tryck **Enter**.
4. Starta Excel. Fliken **Mallverktyg** finns längst till höger i menyfliksområdet. Syns den
   inte: **Infoga → Mina tillägg → Utvecklartillägg** och klicka på Kvanto Mallverktyg.

## Installera

```bat
mkdir "%LOCALAPPDATA%\KvantoMallverktyg" 2>nul & curl -fsSL https://kvanto-mallverktyg.aleks-bok.workers.dev/mallverktyg-manifest.xml -o "%LOCALAPPDATA%\KvantoMallverktyg\manifest.xml" && reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 8c511835-2eb9-4d7e-b149-c617fc608e15 /t REG_SZ /d "%LOCALAPPDATA%\KvantoMallverktyg\manifest.xml" /f
```

Kommandot hämtar Mallverktygets manifest (en liten textfil som talar om för Excel var tillägget
finns) och registrerar det för din användare. Inget program installeras och inga
administratörsrättigheter behövs. Kör kommandot igen om nya knappar har tillkommit.

## Ta bort

```bat
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 8c511835-2eb9-4d7e-b149-c617fc608e15 /f & rmdir /s /q "%LOCALAPPDATA%\KvantoMallverktyg"
```

## Hela byrån på en gång

Har byrån Microsoft 365 kan en administratör lägga till Mallverktyget för alla under
admin.microsoft.com → **Inställningar → Integrerade appar** med den här länken:

```text
https://kvanto-mallverktyg.aleks-bok.workers.dev/mallverktyg-manifest.xml
```

Mer hjälp: https://kvanto-mallverktyg.aleks-bok.workers.dev/installera.html
