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
  
  async getContactGroups(limit: number = 50): Promise<any[]> {
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
          filter: { kind: 'group' },
          limit
        }, 'query'],
        ['ContactCard/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
        }, 'groups']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(`Contact groups not supported or accessible: ${error instanceof Error ? error.message : String(error)}`);
    }
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
    firstName?: string;
    lastName?: string;
    name?: string; // full name override (if provided without first/last, will attempt to split)
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    company?: string;
    notes?: string;
    birthday?: string; // YYYY-MM-DD format
    addresses?: Array<{
      street?: string;
      locality?: string; // city
      region?: string; // state/province
      postcode?: string;
      country?: string;
      countryCode?: string; // ISO 3166-1 (e.g. "US")
      label?: string;
    }>;
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

    // Resolve name components
    let firstName = contact.firstName;
    let lastName = contact.lastName;
    let fullName = contact.name;

    if (!firstName && !lastName && fullName) {
      // Split full name into first/last as best effort
      const parts = fullName.trim().split(/\s+/);
      if (parts.length === 1) {
        firstName = parts[0];
      } else {
        firstName = parts.slice(0, -1).join(' ');
        lastName = parts[parts.length - 1];
      }
    }

    if (!fullName) {
      fullName = [firstName, lastName].filter(Boolean).join(' ');
    }

    // Build name with components (required for Fastmail to display the name)
    const nameComponents: Array<{ kind: string; value: string }> = [];
    if (firstName) nameComponents.push({ kind: 'given', value: firstName });
    if (lastName) nameComponents.push({ kind: 'surname', value: lastName });

    // Build JSContact Card object
    const card: Record<string, any> = {
      '@type': 'Card',
      version: '1.0',
      addressBookIds: { [addressBookId as string]: true },
      name: {
        full: fullName,
        ...(nameComponents.length > 0 && { components: nameComponents }),
      },
    };

    if (contact.emails?.length) {
      card.emails = {};
      contact.emails.forEach((e, i) => {
        const email: Record<string, any> = { address: e.address };
        if (e.label) {
          // Map common labels to JSContact contexts
          const ctx = e.label.toLowerCase();
          if (ctx === 'work') email.contexts = { work: true };
          else email.contexts = { private: true };
        }
        card.emails[`e${i}`] = email;
      });
    }

    if (contact.phones?.length) {
      card.phones = {};
      contact.phones.forEach((p, i) => {
        const phone: Record<string, any> = { number: p.number };
        if (p.label) {
          // Map common labels to JSContact features
          const feat = p.label.toLowerCase();
          if (feat === 'mobile' || feat === 'cell') {
            phone.features = { mobile: true, voice: true };
          } else if (feat === 'fax') {
            phone.features = { fax: true };
          } else if (feat === 'pager') {
            phone.features = { pager: true };
          } else {
            phone.features = { voice: true };
          }
        }
        card.phones[`p${i}`] = phone;
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

    if (contact.addresses?.length) {
      card.addresses = {};
      contact.addresses.forEach((a, i) => {
        const components: Array<{ kind: string; value: string }> = [];
        if (a.street) components.push({ kind: 'name', value: a.street });
        if (a.locality) components.push({ kind: 'locality', value: a.locality });
        if (a.region) components.push({ kind: 'region', value: a.region });
        if (a.postcode) components.push({ kind: 'postcode', value: a.postcode });
        if (a.country) components.push({ kind: 'country', value: a.country });

        const addr: Record<string, any> = {
          components,
          pref: 1,
        };
        if (a.countryCode) addr.countryCode = a.countryCode.toLowerCase();
        if (a.label) {
          const ctx = a.label.toLowerCase();
          addr.contexts = ctx === 'work' ? { work: true } : { private: true };
        }
        card.addresses[`ad${i}`] = addr;
      });
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

  async createContactGroup(group: {
    name: string;
    memberIds?: string[];
    addressBookId?: string;
  }): Promise<string> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();

    // Get address book ID - use provided one or fetch default
    let addressBookId = group.addressBookId;
    if (!addressBookId) {
      const addressBooks = await this.getAddressBooks();
      if (addressBooks.length === 0) {
        throw new Error('No address books found');
      }
      addressBookId = addressBooks[0].id;
    }

    // Build members map
    const members: Record<string, boolean> = {};
    if (group.memberIds?.length) {
      for (const id of group.memberIds) {
        members[id] = true;
      }
    }

    const card: Record<string, any> = {
      '@type': 'Card',
      version: '1.0',
      kind: 'group',
      addressBookIds: { [addressBookId as string]: true },
      name: { full: group.name },
      members,
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId: session.accountId,
          create: { newGroup: card }
        }, 'createGroup']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notCreated?.newGroup) {
        throw new Error(`Failed to create contact group: ${JSON.stringify(result.notCreated.newGroup)}`);
      }
      const groupId = result.created?.newGroup?.id;
      if (!groupId) {
        throw new Error('Contact group creation returned no ID');
      }
      return groupId;
    } catch (error) {
      throw new Error(`Contact group creation failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async updateContact(contactId: string, updates: {
    firstName?: string;
    lastName?: string;
    name?: string;
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    company?: string;
    notes?: string;
    birthday?: string;
    addresses?: Array<{
      street?: string;
      locality?: string;
      region?: string;
      postcode?: string;
      country?: string;
      countryCode?: string;
      label?: string;
    }>;
  }): Promise<void> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();

    // Build patch object - only include fields that are provided
    const patch: Record<string, any> = {};

    // Handle name updates
    if (updates.firstName !== undefined || updates.lastName !== undefined || updates.name !== undefined) {
      let firstName = updates.firstName;
      let lastName = updates.lastName;
      let fullName = updates.name;

      if (!firstName && !lastName && fullName) {
        const parts = fullName.trim().split(/\s+/);
        if (parts.length === 1) {
          firstName = parts[0];
        } else {
          firstName = parts.slice(0, -1).join(' ');
          lastName = parts[parts.length - 1];
        }
      }

      if (!fullName) {
        fullName = [firstName, lastName].filter(Boolean).join(' ');
      }

      const nameComponents: Array<{ kind: string; value: string }> = [];
      if (firstName) nameComponents.push({ kind: 'given', value: firstName });
      if (lastName) nameComponents.push({ kind: 'surname', value: lastName });

      patch['name'] = {
        full: fullName,
        ...(nameComponents.length > 0 && { components: nameComponents }),
      };
    }

    if (updates.emails !== undefined) {
      if (updates.emails.length === 0) {
        patch['emails'] = null;
      } else {
        const emails: Record<string, any> = {};
        updates.emails.forEach((e, i) => {
          const email: Record<string, any> = { address: e.address };
          if (e.label) {
            const ctx = e.label.toLowerCase();
            if (ctx === 'work') email.contexts = { work: true };
            else email.contexts = { private: true };
          }
          emails[`e${i}`] = email;
        });
        patch['emails'] = emails;
      }
    }

    if (updates.phones !== undefined) {
      if (updates.phones.length === 0) {
        patch['phones'] = null;
      } else {
        const phones: Record<string, any> = {};
        updates.phones.forEach((p, i) => {
          const phone: Record<string, any> = { number: p.number };
          if (p.label) {
            const feat = p.label.toLowerCase();
            if (feat === 'mobile' || feat === 'cell') {
              phone.features = { mobile: true, voice: true };
            } else if (feat === 'fax') {
              phone.features = { fax: true };
            } else if (feat === 'pager') {
              phone.features = { pager: true };
            } else {
              phone.features = { voice: true };
            }
          }
          phones[`p${i}`] = phone;
        });
        patch['phones'] = phones;
      }
    }

    if (updates.company !== undefined) {
      patch['organizations'] = updates.company ? { o0: { name: updates.company } } : null;
    }

    if (updates.notes !== undefined) {
      patch['notes'] = updates.notes ? { n0: { note: updates.notes } } : null;
    }

    if (updates.birthday !== undefined) {
      if (updates.birthday) {
        const [year, month, day] = updates.birthday.split('-').map(Number);
        const dateObj: Record<string, number> = {};
        if (year) dateObj.year = year;
        if (month) dateObj.month = month;
        if (day) dateObj.day = day;
        patch['anniversaries'] = { a0: { kind: 'birth', date: dateObj } };
      } else {
        patch['anniversaries'] = null;
      }
    }

    if (updates.addresses !== undefined) {
      if (updates.addresses.length === 0) {
        patch['addresses'] = null;
      } else {
        const addresses: Record<string, any> = {};
        updates.addresses.forEach((a, i) => {
          const components: Array<{ kind: string; value: string }> = [];
          if (a.street) components.push({ kind: 'name', value: a.street });
          if (a.locality) components.push({ kind: 'locality', value: a.locality });
          if (a.region) components.push({ kind: 'region', value: a.region });
          if (a.postcode) components.push({ kind: 'postcode', value: a.postcode });
          if (a.country) components.push({ kind: 'country', value: a.country });

          const addr: Record<string, any> = { components, pref: 1 };
          if (a.countryCode) addr.countryCode = a.countryCode.toLowerCase();
          if (a.label) {
            const ctx = a.label.toLowerCase();
            addr.contexts = ctx === 'work' ? { work: true } : { private: true };
          }
          addresses[`ad${i}`] = addr;
        });
        patch['addresses'] = addresses;
      }
    }

    if (Object.keys(patch).length === 0) {
      throw new Error('No updates provided');
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId: session.accountId,
          update: { [contactId]: patch }
        }, 'updateContact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notUpdated?.[contactId]) {
        throw new Error(`Failed to update contact: ${JSON.stringify(result.notUpdated[contactId])}`);
      }
    } catch (error) {
      throw new Error(`Contact update failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async addContactToGroup(groupId: string, contactUids: string[]): Promise<void> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();

    // Use JMAP patch syntax to add members (keyed by contact UID)
    const patch: Record<string, any> = {};
    for (const uid of contactUids) {
      patch[`members/${uid}`] = true;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId: session.accountId,
          update: { [groupId]: patch }
        }, 'addToGroup']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notUpdated?.[groupId]) {
        throw new Error(`Failed to add contacts to group: ${JSON.stringify(result.notUpdated[groupId])}`);
      }
    } catch (error) {
      throw new Error(`Add to group failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async removeContactFromGroup(groupId: string, contactUids: string[]): Promise<void> {
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled.');
    }

    const session = await this.getSession();

    // Use JMAP patch syntax to remove members (set to null, keyed by contact UID)
    const patch: Record<string, any> = {};
    for (const uid of contactUids) {
      patch[`members/${uid}`] = null;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId: session.accountId,
          update: { [groupId]: patch }
        }, 'removeFromGroup']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      if (result.notUpdated?.[groupId]) {
        throw new Error(`Failed to remove contacts from group: ${JSON.stringify(result.notUpdated[groupId])}`);
      }
    } catch (error) {
      throw new Error(`Remove from group failed: ${error instanceof Error ? error.message : String(error)}`);
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