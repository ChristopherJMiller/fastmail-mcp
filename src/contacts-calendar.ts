import { JmapClient, JmapRequest } from './jmap-client.js';

export class ContactsCalendarClient extends JmapClient {
  
  private async checkContactsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:contacts'];
  }
  
  private async checkCalendarsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:calendars'];
  }
  
  async getContacts(limit: number = 50): Promise<any[]> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/query', {
          accountId: session.accountId,
          limit
        }, 'query'],
        ['ContactCard/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async getContactById(id: string): Promise<any> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 0)[0];
    } catch (error) {
      throw new Error(`Contact access not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async searchContacts(query: string, limit: number = 20): Promise<any[]> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/query', {
          accountId: session.accountId,
          filter: { text: query },
          limit
        }, 'query'],
        ['ContactCard/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(`Contact search not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async getCalendars(): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['Calendar/get', {
          accountId: session.accountId
        }, 'calendars']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 0);
    } catch (error) {
      // Calendar access might require special permissions
      throw new Error(`Calendar access not supported or requires additional permissions. This may be due to account settings or JMAP scope limitations: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();
    
    const filter = calendarId ? { inCalendar: calendarId } : {};
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'start', isAscending: true }],
          limit
        }, 'query'],
        ['CalendarEvent/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'CalendarEvent/query', path: '/ids' },
          properties: ['id', 'title', 'description', 'start', 'end', 'location', 'participants']
        }, 'events']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(`Calendar events access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEventById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'event']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 0)[0];
    } catch (error) {
      throw new Error(`Calendar event access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getAddressBooks(): Promise<any[]> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['AddressBook/get', {
          accountId: session.accountId
        }, 'addressbooks']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0);
  }

  async createContact(contact: {
    name: string;
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    company?: string;
    notes?: string;
    birthday?: string; // YYYY-MM-DD format
    addressBookId?: string;
  }): Promise<string> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();

    // Get address book ID - use provided one or fetch default
    let addressBookId = contact.addressBookId;
    if (!addressBookId) {
      const addressBooks = await this.getAddressBooks();
      if (addressBooks.length === 0) {
        throw new Error('No address books found');
      }
      addressBookId = addressBooks[0].id;
    }

    // Build JSContact Card object
    const card: Record<string, any> = {
      '@type': 'Card',
      version: '1.0',
      addressBookIds: { [addressBookId as string]: true },
      name: { full: contact.name },
    };

    if (contact.emails?.length) {
      card.emails = {};
      contact.emails.forEach((e, i) => {
        card.emails[`e${i}`] = { address: e.address, ...(e.label && { label: e.label }) };
      });
    }

    if (contact.phones?.length) {
      card.phones = {};
      contact.phones.forEach((p, i) => {
        card.phones[`p${i}`] = { number: p.number, ...(p.label && { label: p.label }) };
      });
    }

    if (contact.company) {
      card.organizations = { o0: { name: contact.company } };
    }

    if (contact.notes) {
      card.notes = { n0: { note: contact.notes } };
    }

    if (contact.birthday) {
      const [year, month, day] = contact.birthday.split('-').map(Number);
      const dateObj: Record<string, number> = {};
      if (year) dateObj.year = year;
      if (month) dateObj.month = month;
      if (day) dateObj.day = day;
      card.anniversaries = { a0: { kind: 'birth', date: dateObj } };
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId: session.accountId,
          create: { newContact: card }
        }, 'createContact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notCreated?.newContact) {
        throw new Error(`Failed to create contact: ${JSON.stringify(result.notCreated.newContact)}`);
      }
      const contactId = result.created?.newContact?.id;
      if (!contactId) {
        throw new Error('Contact creation returned no ID');
      }
      return contactId;
    } catch (error) {
      throw new Error(`Contact creation failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async deleteContact(contactId: string): Promise<void> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId: session.accountId,
          destroy: [contactId]
        }, 'deleteContact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notDestroyed?.[contactId]) {
        throw new Error(`Failed to delete contact: ${JSON.stringify(result.notDestroyed[contactId])}`);
      }
    } catch (error) {
      throw new Error(`Contact deletion failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string; // ISO 8601 format
    end: string;   // ISO 8601 format
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
  }): Promise<string> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    const eventObject = {
      calendarId: event.calendarId,
      title: event.title,
      description: event.description || '',
      start: event.start,
      end: event.end,
      location: event.location || '',
      participants: event.participants || []
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          create: { newEvent: eventObject }
        }, 'createEvent']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      const eventId = result.created?.newEvent?.id;
      if (!eventId) {
        throw new Error('Calendar event creation returned no event ID');
      }
      return eventId;
    } catch (error) {
      throw new Error(`Calendar event creation not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }
}