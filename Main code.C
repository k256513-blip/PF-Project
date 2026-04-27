#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <windows.h> 
#include <conio.h> 

struct TheftRecord {
    int id;
    char date[20];
    char time[20];
    char location[50];
};

struct Owner {
    char name[50];
    int bicycleID;
    char phone[15];
};

// --- MODIFIED: LOG TO EXCEL-COMPATIBLE CSV ---
void autoLogTheft() {
    // Open in "append" mode. Creates the file if it doesn't exist.
    FILE *fp = fopen("Theft_History.csv", "a");
    if (fp == NULL) return;

    // Check if file is empty to add headers
    fseek(fp, 0, SEEK_END);
    if (ftell(fp) == 0) {
        fprintf(fp, "Record ID,Date,Time,Location\n");
    }

    int id = rand() % 9000 + 1000;
    char dateStr[20], timeStr[20];

    time_t t = time(NULL);
    struct tm *tm_info = localtime(&t);
    strftime(dateStr, 20, "%d-%m-%Y", tm_info);
    strftime(timeStr, 20, "%H:%M:%S", tm_info);

    // Write data in Comma Separated format
    fprintf(fp, "%d,%s,%s,Bicycle Stand A\n", id, dateStr, timeStr);
    fclose(fp);

    printf("\n[DATABASE] Recorded to Excel: ID %d at %s\n", id, timeStr);
}

// --- MODIFIED: OWNER REGISTER TO CSV ---
void addOwner() {
    FILE *fp = fopen("Owner_Database.csv", "a");
    struct Owner o;
    if (fp == NULL) return;

    fseek(fp, 0, SEEK_END);
    if (ftell(fp) == 0) {
        fprintf(fp, "Owner Name,Bicycle ID,Phone Number\n");
    }
    
    printf("\nEnter Owner Name: "); scanf(" %[^\n]", o.name);
    printf("Enter Bicycle ID: "); scanf("%d", &o.bicycleID);
    printf("Enter Phone: "); scanf("%s", o.phone);
    
    fprintf(fp, "%s,%d,%s\n", o.name, o.bicycleID, o.phone);
    fclose(fp);
    printf("Owner saved to Excel database.\n");
}

void resetHistory() {
    char confirm;
    printf("\nDelete all Excel records? (y/n): ");
    scanf(" %c", &confirm);
    if (confirm == 'y' || confirm == 'Y') {
        remove("Theft_History.csv");
        printf("Excel history file deleted.\n");
    }
}

// --- MONITOR LOGIC ---
void startMonitor() {
    HANDLE hSerial;
    DCB dcb = {0};
    const char* port = "\\\\.\\COM11"; 

    hSerial = CreateFile(port, GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (hSerial == INVALID_HANDLE_VALUE) {
        printf("\n[ERROR] Port COM11 busy or not found.\n");
        return;
    }

    dcb.DCBlength = sizeof(dcb);
    GetCommState(hSerial, &dcb);
    dcb.BaudRate = CBR_9600;
    dcb.ByteSize = 8;
    dcb.StopBits = ONESTOPBIT;
    dcb.Parity   = NOPARITY;
    SetCommState(hSerial, &dcb);

    printf("\n--- MONITORING (ESC to Exit) ---\n");

    char buffer[256];
    DWORD bytesRead;

    while (1) {
        if (kbhit()) { if (getch() == 27) break; }

        if (ReadFile(hSerial, buffer, sizeof(buffer) - 1, &bytesRead, NULL) && bytesRead > 0) {
            buffer[bytesRead] = '\0';

            if (strstr(buffer, "THEFT_DETECTED") != NULL) {
                printf("\n>>> SENSOR TRIGGERED! <<<");
                autoLogTheft();
            }
            if (strstr(buffer, "STOPPED_BY_BUTTON") != NULL) {
                printf("[STATUS] Alarm deactivated by user.\n");
            }
            if (strstr(buffer, "ALARM_TIMEOUT") != NULL) {
                printf("[STATUS] Alarm timed out.\n");
            }
        }
        Sleep(50); 
    }
    CloseHandle(hSerial);
}

int main() {
    int choice;
    srand(time(NULL));

    while (1) {
        printf("\n--- BICYCLE SECURITY (EXCEL LOGGING) ---");
        printf("\n1. START LIVE MONITOR (Arduino)");
        printf("\n2. Register New Owner (Excel)");
        printf("\n3. RESET EXCEL HISTORY");
        printf("\n4. Exit");
        printf("\nSelect: ");
        
        scanf("%d", &choice);
        switch (choice) {
            case 1: startMonitor(); break; 
            case 2: addOwner(); break;
            case 3: resetHistory(); break;
            case 4: exit(0);
            default: printf("Invalid option.\n");
        }
    }
    return 0;
}
