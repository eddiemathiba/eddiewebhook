'use strict';

// Import the Dialogflow module from Google client libraries.
const functions = require('firebase-functions');
const { google } = require('googleapis');
const { WebhookClient } = require('dialogflow-fulfillment');

// Enter your calendar ID below and service account JSON below
// Enter your calendar ID below and service account JSON below
const calendarId = "key here";
const serviceAccount = {
    "type": "service_account",
    "project_id": "libbyonline-jucplk",
    "private_key_id": "key here",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCg7M4iUN+xgeYz\ndepvGwgEdZ/7jpV4PxBmpMLZS2FeuhHZ4ZasdTeif1HUUiamFZuRvXsWzrdH2Rp5\nTkX2nFTszaKaM8eqSEblRoTYvsUq3suBwPYH3Rf6JyGKJiv/gM4DXUJ0BCCb6pGV\nbMh6IBiVFSa0s62OIcdelfCskeSsdbDEWkOEsoKt6sEQ1IyLcLBe7jREIoanf3Js\ncQf6oIdSoAGdKn+67dj4LggR5JGMNF/qSFbuWtJhb8nm66RLtLUQYgu0Udzd4Hmi\nJNktY7ZJWFje4BJrlTn/dbbaO3dcUxJdmq3Dzk5J1VKL5k9l+v7b522xVTNOee7U\n7aoMyEhxAgMBAAECggEAAdXhngOVIZWtNf4Mz/qbc02SJxDfxJDXk4wiis7wy99M\nb9NCYDIwEcLpAIZ1WWSasyVZn495CLFBu4t4gdETqOXJin+3scoEQP42arB2WEBV\nSuQUCk/rw5cpq0U3MEFcWC56oQ8g/hZFVdJ5UOdH0H9+wwXfn2TkPMByD83UMWFb\nolUUfz/j+G+e9qTCnUIurUod0u4s4tggQmYGB/iH0lQLRoeUiCPwA5ixhfhTzkp6\nceOhF1ZPtlbsugKKzf+LJVI1ew9olBE2mnyWfw82tK96tZmSLwZKSlDHtBxyLj+C\nUWrytqfNXZNGwFs3gsaJZnR4zkWKs1lPVaVZ92WmtwKBgQDRw8s7qc7MzzH5cY+A\nk2w0vl2s7hM/xdzjOH4278VjsUnskK9XOJ+puplXFm6O3ua4J5SCFSOchqZKOpko\nYe0ws4InMKgBWaBOtMljlT/RHzp4Rwvr6SqBFBvbLjJmMtc/ueb+zK62vJaJL6BG\nAnpS8BLPQG/D84es/OA3rp/30wKBgQDEZSxI+846+x2sTgbMmsNDeJDVD9pBkPpe\nZE/SleE0HgsLbynMpnJvGbicNOBs+ZP0AbFFBUHPZyeoG4jUbdr9Ia0aXKjIWKgM\n5GMrvZG9ISYh+olkzqBiENPxoK65l93rUp2MgtkI8/BrTj40iwrMzTkLx0CK9kFO\nXANOJs64KwKBgA9ixBppNyDwpaK9QKBWr94ihG51u/W2hqAn+aH/7kOEsn5mkmXc\nYmGprSulGpDiKpwhUxRDhqwpeDMDn05F+IbU89v6BqiqsUZ0njzuqpKlLr25wFca\ncCKtnelytwGmAwHKmfXgf4QpywTe+GuKLPy+XTLUcB44U6BNoAiVh2g/AoGBALAx\nAADd/s+FN8K8IXdvpJwkTvMGjMzjrun93mrTUb268jeo2/wgm2s2zRs+iNTbNzPi\nTNHZ/IeASOCVCzVg9OTBUJXk6PsMJn+iaaH1KQle8uHq7VYF1XcFF8+fUwFn0Izs\nijFjtihFCLyL7lyfHOpNds8tS6cfc8lD3SdAw5YPAoGAU4XaCo1J+oP1smqKH7Io\noHhZJ9j0ELbleOjrWzybSduYhC3BcDoUSM355wFs78AAR9ocPIHkzw7KQcKqzN6O\nZyRDibwa7BkW+GX0KW85K09p4yUv+MJI7A8tzcf9P43zK6FlCZO2pUdC8M8g/tpQ\nDSNjaxUuFyYrlNhEjPWgi9s=\n-----END PRIVATE KEY-----\n",
    "client_email": "key here",
    "client_id": "114392323486401614774",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": ""}; // Starts with {"type": "service_account",...

// Set up Google Calendar Service account credentials
const serviceAccountAuth = new google.auth.JWT({
    email: serviceAccount.client_email,
    key: serviceAccount.private_key,
    scopes: 'https://www.googleapis.com/auth/calendar'
});

const calendar = google.calendar('v3');
process.env.DEBUG = 'dialogflow:*'; // enables lib debugging statements

const timeZone = 'South Africa standard time';
const timeZoneOffset = 'GMT+02:00';

// Set the DialogflowApp object to handle the HTTPS POST request.
exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
    const agent = new WebhookClient({ request, response });
    console.log("Parameters", agent.parameters);
    const appointment_type = agent.parameters.AppointmentType;

    function makeAppointment(agent) {
        // Calculate appointment start and end datetimes (end = +1hr from start)
        const dateTimeStart = new Date(new Date(Date.parse(agent.parameters.date.split('T')[0] + 'T' + agent.parameters.time.split('T')[1].split('-')[0])));
        console.log("expected String", agent.parameters.date.split('T')[0] + 'T' + agent.parameters.time.split('T')[1].split('-')[0] + timeZoneOffset);
        const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
        const appointmentTimeString = dateTimeStart.toLocaleString(
            'en-US', { month: 'long', day: 'numeric', hour: 'numeric', timeZone: timeZone }
        );
        // Check the availability of the time, and make an appointment if there is time on the calendar
        console.log("dateTimeStart", dateTimeStart);
        console.log("dateTimeEnd", dateTimeEnd);
        console.log("appointmentTimeString", appointmentTimeString);
        return createCalendarEvent(dateTimeStart, dateTimeEnd, appointment_type).then(() => {
            agent.add(`Ok, let me see if we can fit you in. ${appointmentTimeString} is fine!.`);
        }).catch(() => {
            agent.add(`I'm sorry, there are no slots available for ${appointmentTimeString}.`);
        });
    }

    // Handle the Dialogflow intent named 'Schedule Appointment'.
    let intentMap = new Map();
    intentMap.set('schedule appointment', makeAppointment);
    agent.handleRequest(intentMap);
});

//Creates calendar event in Google Calendar
function createCalendarEvent(dateTimeStart, dateTimeEnd, appointment_type) {
    return new Promise((resolve, reject) => {
        calendar.events.list({
            auth: serviceAccountAuth, // List events for time period
            calendarId: calendarId,
            timeMin: dateTimeStart.toISOString(),
            timeMax: dateTimeEnd.toISOString()
        }, (err, calendarResponse) => {
            // Check if there is a event already on the Calendar
            if (err || calendarResponse.data.items.length > 0) {
                reject(err || new Error('Requested time conflicts with another appointment'));
            } else {
                // Create event for the requested time period
                calendar.events.insert({
                    auth: serviceAccountAuth,
                    calendarId: calendarId,
                    resource: {
                        summary: appointment_type + ' Appointment',
                        description: appointment_type,
                        start: { dateTime: dateTimeStart },
                        end: { dateTime: dateTimeEnd }
                    }
                }, (err, event) => {
                    err ? reject(err) : resolve(event);
                });
            }
        });
    });
}
