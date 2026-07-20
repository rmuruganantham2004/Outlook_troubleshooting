export declare function readCalendarEvents(top?: number): Promise<any>;
export declare function createCalendarEvent(subject: string, start: string, end: string, timeZone?: string, location?: string, body?: string): Promise<any>;
export declare function deleteCalendarEvent(id: string): Promise<{
    success: boolean;
}>;
