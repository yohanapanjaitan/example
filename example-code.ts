import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { google, calendar_v3 } from 'googleapis';
import * as moment from "moment"
import { GoogleCalendarRequestDto } from '../dto/google-calendar-request.dto';
import { EventNames } from '../../event-names';
import { CalendarAttendeeUpdatedEvent } from '../../calendars/events/CalendarAttendeeUpdatedEvent';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { AttendeeType } from '../enum/attendee-type-enum';
import { GoogleCalendarAttendeeDto } from '../dto/google-calendar-attendee.dto';
import { GoogleCalendarDescriptionDto } from '../dto/google-calendar-description.dto';
import { Program } from '../../program/entities/programs.entity';
import { Training } from '../../trainings/trainings.entity';
import { CryptUtil } from '../../common/util/crypt.util';
import * as aixpTracker from 'aixp-node-sdk';
import { v4 as uuid } from 'uuid';
import { ProgramType } from '../../back-office/program/enum/program-type.enum';
import { ProgramFormat } from '../../back-office/program/enum/program-format.enum';
import { GoogleCalendarListRequestDto } from '../dto/google-calendar-list-request.dto';
import { RskCacheService } from '../../rsk-cache/rsk-cache.service';
import { GoogleCalendarCacheValueDto } from '../dto/google-calendar-cache-value.dto';
import { GoogleCalendarResultDto } from '../dto/google-calendar-result.dto';
import { ProductType } from '../../trainings/product-type.enum';
import { GoogleCalendarUpdateResultDto } from '../dto/google-calendar-update-result.dto';

@Injectable()
export class GoogleCalendarService {
  private scopes: string[];
  private keyFilePath: string;
  private webHost: string;
  private encryptKey: string;
  private readonly aixpLoadId: string;
  private readonly aixpUrl: string;
  private prakerjaOrganizers: string[];
  private nonPrakerjaOrganizers: string[];
  private externalAttendeesLimit: number;
  private gcalMainKey: string;
  private gcalBlockKey: string;
  constructor(
    private readonly configService: ConfigService,
    private readonly eventEmitter: EventEmitter2,
    private readonly rskCacheService: RskCacheService
  ) {
    this.keyFilePath = '\google-service-account.json'
    this.scopes = [
      'https://www.googleapis.com/auth/calendar',
      'https://www.googleapis.com/auth/calendar.events',
      'https://www.googleapis.com/auth/admin.directory.resource.calendar '
    ];
    this.webHost = this.configService.get<string>('WEB_HOST')
    this.encryptKey = this.configService.get<string>('ENCRYPT_KEY_FOR_GCAL_LINK_TRACKER')
    this.aixpLoadId = this.configService.get<string>('AIXP_LOAD_ID')
    this.aixpUrl = this.configService.get<string>('AIXP_URL')
    this.prakerjaOrganizers = this.configService.get<string>('GCAL_PRAKERJA_ORGS').split(',')
    this.nonPrakerjaOrganizers = this.configService.get<string>('GCAL_NON_PRAKERJA_ORGS').split(',')
    this.externalAttendeesLimit = +this.configService.get<number>('GCAL_EXT_ATTENDEE_LIMIT', 2000)
    this.gcalMainKey = 'google-calendar'
    this.gcalBlockKey = 'google-calendar-blocked-organizer'
  }

  async authorize(clientEmail: string = null): Promise<calendar_v3.Calendar> {
    const auth = new google.auth.GoogleAuth({
      keyFile: this.keyFilePath,
      scopes: this.scopes,
      clientOptions: {
        subject: clientEmail
      }
    });
    return google.calendar({
      version: 'v3',
      auth
    });
  }

  async fetchCalendarList(): Promise<any> {
    const googleCalendar = await this.authorize()
    const result = await googleCalendar.calendarList.get({
      calendarId: 'primary', // 'primary' represents the primary calendar of the authenticated user
    })
    return result.data
  }

  async fetchEvents(payload: GoogleCalendarListRequestDto, organizer: string): Promise<any> {
    try {
      const googleCalendar = await this.authorize(organizer)
      const calendarPrimary = await googleCalendar.events.list({
        calendarId: 'primary', // 'primary' represents the primary calendar of the authenticated user
        timeZone: payload.timeZone || "Asia/Jakarta",
        timeMax: payload.timeMax || undefined, // example timeMax: "2023-07-21T10:00:00+07:00"
        timeMin: payload.timeMin || undefined,
      })
      const result = calendarPrimary.data
      return result
    } catch (err) {
      console.error('Error fetching calendar event list', err);
    }
  }

  async fetchOneEvent(eventId: string, organizer: string): Promise<any> {
    try {
      const googleCalendar = await this.authorize(organizer)
      const calendarPrimary = await googleCalendar.events.get({
        calendarId: 'primary', // 'primary' represents the primary calendar of the authenticated user
        eventId
      })
      const result = calendarPrimary.data
      return result

    } catch (err) {
      console.error('Error fetching calendar event', err);
    }
  }

  async fetchCacheData(sufixKey: string): Promise<any> {
    const cacheKey = `${this.gcalMainKey}-${sufixKey}`
    const data = await this.rskCacheService.get(cacheKey)
    return this.deserializeCacheValue(data)
  }

  async deleteEvent(eventId: string, organizer: string = null): Promise<any> {
    try {
      const googleCalendar = await this.authorize(organizer)
      const result = await googleCalendar.events.delete({
        calendarId: 'primary', // 'primary' represents the primary calendar of the authenticated user
        eventId
      })
      console.log('Successfully deleted google calendar event id', eventId);
      return result
    } catch (err) {
      console.log('Error deleting calendar event', err);
    }
  }

  async createEvent(createCalendar?: GoogleCalendarRequestDto, organizer: string = null): Promise<any> {
    const { eventName, attendees, startTime, endTime, programName, location, conferenceLink, guestsCanSeeOtherGuests, training, program, isYoutubeLiveIdExist } = createCalendar
    const description = this.descriptionConstructor(
      programName,
      startTime,
      endTime,
      conferenceLink ? conferenceLink : '',
      location,
      training,
      program,
      isYoutubeLiveIdExist
    )
    const studentAttendee = attendees.find((attendee) => attendee.type === AttendeeType.STUDENT);

    const event = {
      summary: eventName,
      start: {
        dateTime: startTime,
        timeZone: "Asia/Jakarta",
      },
      end: {
        dateTime: endTime,
        timeZone: "Asia/Jakarta",
      },
      attendees,
      description,
      location,
      guestsCanSeeOtherGuests
    };

    try {
      const isPrakerja = training?.productType === ProductType.PRAKERJA || program?.productType === ProductType.PRAKERJA
      const currentOrganizerCached = (await this.fetchCurrentOrganizer(isPrakerja))?.organizer
      const isAvailable = await this.isOrganizerAvailable(currentOrganizerCached)
      if (!organizer) {
        if (currentOrganizerCached && isAvailable) {
          organizer = currentOrganizerCached
        } else {
          const organizers = await this.fetchAvailableOrganizers(isPrakerja)
          if (organizers && organizers[0]) {
            organizer = organizers[0]
          } else {
          }
        }
      }
      const googleCalendar = await this.authorize(organizer)
      const response = await googleCalendar.events.insert({
        calendarId: 'primary',
        requestBody: event
      });
      if (response.status == 200) {
        console.log(`Successfully created google calendar event organizer ${organizer} for user`, attendees)
        const client = new aixpTracker(`${this.aixpLoadId}`, this.aixpUrl)
        client.track({
          anonymousId: uuid(),
          event: "addGoogleCalendarEvent",
          properties: {
            programName: training ? training.title : program.name,
            productType: training ? training.deliveryType : this.defineProgramType(program),
            programType: training ? training.productType : program.productType
          }
        })
        organizer = await this.cacheOrganizerAndAttendee(attendees, organizer, isPrakerja)
      }

      this.eventEmitter.emit(
        EventNames.CALENDAR_ATTENDEE_UPDATED,
        new CalendarAttendeeUpdatedEvent({
          email: studentAttendee.email,
          programName
        }));

      return GoogleCalendarResultDto.fromPayload(response.data, organizer)
    } catch (error) {
      console.error('Error creating google calendar event:', error);
      if (error.message.includes('Calendar usage limit')) {
        await this.blockOrganizer(organizer)
        throw new HttpException('Calendar usage limit', HttpStatus.BAD_REQUEST);
      }
    }
  }

  async updateEventAttendee(
    eventId: string,
    newAttendees: GoogleCalendarAttendeeDto[],
    programName: string,
    training: Training = null,
    program: Program = null,
    organizer: string = null
  ): Promise<any> {
    try {
      const googleCalendar = await this.authorize(organizer)
      const event = await googleCalendar.events.get({
        calendarId: 'primary', // 'primary' represents the primary calendar of the authenticated user
        eventId: eventId,
      });

      let response
      if (newAttendees[0]) {
        const filteredNewAttendees = newAttendees.filter((attendee) => !event.data.attendees.some((oldAttendee) => oldAttendee.email === attendee.email));
        if (filteredNewAttendees[0]) {
          event.data.attendees = [...event.data.attendees, ...filteredNewAttendees]
          response = await googleCalendar.events.update({
            calendarId: 'primary',
            eventId: eventId,
            requestBody: event.data
          });

          /* 
            TO SEND NOTIFICATION, only support one student
          */
          const studentAttendee = newAttendees.find((attendee) => attendee.type === AttendeeType.STUDENT);
          this.eventEmitter.emit(
            EventNames.CALENDAR_ATTENDEE_UPDATED,
            new CalendarAttendeeUpdatedEvent({
              email: studentAttendee.email,
              programName
            }));

          if (response.status == 200) {
            console.log(`Successfully added attendees to google calendar event ${eventId} organizer ${organizer}`, filteredNewAttendees);
            const client = new aixpTracker(`${this.aixpLoadId}`, this.aixpUrl)
            client.track({
              anonymousId: uuid(),
              event: "addAttendeeGoogleCalendarEvent",
              properties: {
                userEmail: filteredNewAttendees[0].email,
                programName: training ? training.title : program.name,
                productType: training ? training.deliveryType : this.defineProgramType(program),
                programType: training ? training.productType : program.productType
              }
            })
            const isPrakerja = training?.productType === ProductType.PRAKERJA || program?.productType === ProductType.PRAKERJA
            await this.cacheOrganizerAndAttendee(newAttendees, organizer, isPrakerja)
          }
        }
      }
      return GoogleCalendarUpdateResultDto.fromPayload(true, null)
    } catch (error) {
      console.error('Error updating google calendar event:', error);
      if (error.message.includes('Calendar usage limit')) {
        await this.blockOrganizer(organizer)
        return GoogleCalendarUpdateResultDto.fromPayload(false, 'Calendar usage limit')
      } else {
        return GoogleCalendarUpdateResultDto.fromPayload(false, error.message)
      }
    }

  }

  /*
    This function is intended for data updating from BO / admin side
    This function cannot update guestsCanSeeOtherGuests & attendees (only possible at order's related function)
  */
  async updateEvent(eventId: string, payload: GoogleCalendarRequestDto, organizer: string = null): Promise<any> {
    /* 
    *  CAUTION !!!
    *  Sending "attendees" payload will overwrite all existing value
    */
    const { eventName, attendees, startTime, endTime, programName, location, conferenceLink, training, program, isYoutubeLiveIdExist } = payload
    try {
      const googleCalendar = await this.authorize(organizer)
      const event = await googleCalendar.events.get({
        calendarId: 'primary', // 'primary' represents the primary calendar of the authenticated user
        eventId: eventId,
      })

      const descriptionConstructionData = GoogleCalendarDescriptionDto.fromPayload({
        description: event.data.description,
        programName,
        startDateTime: startTime,
        endDateTime: endTime,
        conferenceLink,
        location,
        training,
        program,
        isYoutubeLiveIdExist
      })
      event.data.description = this.descriptionPartialConstructor(descriptionConstructionData)

      if (eventName) {
        event.data.summary = eventName
      }
      if (attendees) {
        event.data.attendees = attendees
      }
      if (startTime) {
        event.data.start = {
          dateTime: startTime,
          timeZone: "Asia/Jakarta",
        }
      }
      if (endTime) {
        event.data.end = {
          dateTime: endTime,
          timeZone: "Asia/Jakarta",
        }
      }
      if (location) {
        event.data.location = location
      }
      if (conferenceLink) {
        event.data.location = null
      }

      const updatedEvent = await googleCalendar.events.update({
        calendarId: 'primary',
        eventId: eventId,
        requestBody: event.data
      });
      if (updatedEvent.status == 200) {
        console.log('Successfully updated google calendar event id', eventId);
      }
      return GoogleCalendarUpdateResultDto.fromPayload(true, null)
    } catch (error) {
      console.error('Error updating event:', error);
      if (error.message.includes('Not Found')) {
        return GoogleCalendarUpdateResultDto.fromPayload(false, 'Not Found')
      } else {
        return GoogleCalendarUpdateResultDto.fromPayload(false, error.message)
      }
    }
  }


  private async cacheOrganizerAndAttendee(attendees: GoogleCalendarAttendeeDto[], organizerReq: string = null, isPrakerja: boolean = true): Promise<string> {
    const keyType = isPrakerja ? 'prakerja' : 'non-prakerja'
    const cacheKey = `${this.gcalMainKey}-${keyType}`
    const data = await this.rskCacheService.get(cacheKey)
    const deserializedData = this.deserializeCacheValue(data)
    let organizer: string, uniqueAttendees: string[]
    if (deserializedData) {
      organizer = deserializedData.organizer
      uniqueAttendees = deserializedData.uniqueAttendees
      if (organizer && organizer == organizerReq && uniqueAttendees.length < this.externalAttendeesLimit) {
        const uniqueAttendeesSet = new Set(uniqueAttendees);
        attendees.forEach((attendee) => {
          // Add the email to uniqueAttendees if it's outside RSK domain
          if (!attendee.email.includes('sempurna.com')) {
            uniqueAttendeesSet.add(attendee.email);
          }
        });
        // Convert the Set back to an array, if needed
        uniqueAttendees = Array.from(uniqueAttendeesSet);
      }

    } else {
      organizer = organizerReq
      const newAttendees = []
      for (const att of attendees) {
        if (!att.email.includes('sempurna.com')) newAttendees.push(att.email)
      }
      uniqueAttendees = newAttendees
    }
    const cacheData = GoogleCalendarCacheValueDto.fromPayload(organizer, uniqueAttendees)
    const cacheValue = this.serializeCacheValue(cacheData)
    const oneMonthInMilliSeconds = 30 * 24 * 60 * 60
    await this.rskCacheService.put(cacheKey, cacheValue, oneMonthInMilliSeconds)
    return organizer

  }

  async fetchCurrentOrganizer(isPrakerja: boolean = true): Promise<GoogleCalendarCacheValueDto> {
    const sufixKey = isPrakerja ? 'prakerja' : 'non-prakerja'
    const cachekey = `${this.gcalMainKey}-${sufixKey}`
    const googleCalendarCacheValue = await this.rskCacheService.get(cachekey)
    if (googleCalendarCacheValue) {
      const googleCalendarCacheValueParsed = this.deserializeCacheValue(googleCalendarCacheValue)
      return googleCalendarCacheValueParsed
    }
  }

  async fetchAvailableOrganizers(isPrakerja: boolean = true, organizers: string[] = null): Promise<string[]> {
    const availableOrganizers: string[] = []
    if (!organizers) {
      if (isPrakerja) {
        organizers = this.prakerjaOrganizers
      }
      else {
        organizers = this.nonPrakerjaOrganizers
      }
    }
    for await (const org of organizers) {
      const isAvailable = await this.isOrganizerAvailable(org)
      if (isAvailable) availableOrganizers.push(org)
    }
    return availableOrganizers
  }

  async isOrganizerAvailable(organizer: string): Promise<boolean> {
    // validate if organizer is same as current organizer and if the limit is reached
    const currentOrganizerPrakerja = await this.fetchCurrentOrganizer(true)
    const currentOrganizerNonPrakerja = await this.fetchCurrentOrganizer(false)
    // validate for prakerja org
    const shouldBlockCurrentOrganizerPrakerja = currentOrganizerPrakerja && currentOrganizerPrakerja.organizer === organizer
      && currentOrganizerPrakerja.uniqueAttendees.length >= this.externalAttendeesLimit
    // validate for non-prakerja org
    const shouldBlockCurrentOrganizerNonPrakerja = currentOrganizerNonPrakerja && currentOrganizerNonPrakerja.organizer === organizer
      && currentOrganizerNonPrakerja.uniqueAttendees.length >= this.externalAttendeesLimit

    if (shouldBlockCurrentOrganizerPrakerja || shouldBlockCurrentOrganizerNonPrakerja) {
      if (!shouldBlockCurrentOrganizerPrakerja) await this.blockOrganizer(organizer)
      if (!shouldBlockCurrentOrganizerNonPrakerja) await this.blockOrganizer(organizer)
      return false
    }
    /* 
    *   We need to manually validate blocked date
    *   TTL of redis will be paused if the server's instance is off
    *   Thus, we can't fully rely on redis TTL
    */
    const key = `${this.gcalBlockKey}:${organizer}`
    const cacheValue = await this.rskCacheService.get(key)
    if (cacheValue) {
      const blockedDateStr = await this.deserializeCacheValue(cacheValue)
      const blockedDate = new Date(blockedDateStr)

      const isBlockedTimeFinished = this.isBlockedTimeFinished(blockedDate)
      if (isBlockedTimeFinished) {
        console.log('Limitiation is lifted for google calendar organizer', organizer);
        await this.rskCacheService.delete(key)

        // if the blocked time is finished, the organizer is available (limitation is lifted)
        return true
      }
      return false
    }
    return true
  }

  async blockOrganizer(organizer: string): Promise<any> {
    const key = `${this.gcalBlockKey}:${organizer}`
    const value = await this.serializeCacheValue(new Date())
    const ttl = 24 * 60 * 60 // 24 hours in seconds
    await this.rskCacheService.put(key, value, ttl)
    console.log('Google calendar organizer is blocked for email', organizer);

    const currentOrganizerPrakerja = (await this.fetchCurrentOrganizer())?.organizer
    const currentOrganizerNonPrakerja = await (await this.fetchCurrentOrganizer(false))?.organizer
    if (currentOrganizerPrakerja && organizer === currentOrganizerPrakerja) {
      await this.rskCacheService.delete(`${this.gcalMainKey}-prakerja`)
    } else if (currentOrganizerNonPrakerja && organizer === currentOrganizerNonPrakerja) {
      await this.rskCacheService.delete(`${this.gcalMainKey}-non-prakerja`)
    }

  }

  private descriptionConstructor(
    programTitle: string,
    startDateTime: string,
    endDateTime: string,
    conferenceLink?: string,
    location?: string,
    training?: Training,
    program?: Program,
    isYoutubeLiveIdExist?: boolean,
  ): string {
    const date = moment(startDateTime).format('YYYY-MM-DD')
    const startTime = moment(startDateTime).format('HH:mm')
    const endTime = moment(endDateTime).format('HH:mm')
    const encrypted = this.encryptTrainingOrProgram(training, program)
    const activityPageLink = `${this.webHost}/kelasku?td=${encrypted}`

    // isYoutubeLiveIdExist is used to distinguish, whether the words from ongoingWording use zoom or not
    // if isYoutubeLiveIdExist is true, conferenceLink will set to empty string
    let ongoingWording = '';
    if (isYoutubeLiveIdExist) {
      ongoingWording = 'Sesi ini akan berlangsung';
      conferenceLink = '';
    } else {
      ongoingWording = 'Sesi ini akan berlangsung melalui Zoom';
    }

    if (location) return `Selamat datang di ${programTitle}!\n\nSesi ini akan berlangsung secara luring pada tanggal ${date} pukul ${startTime} hingga ${endTime} WIB. Tempat: ${location}\n\nUntuk melihat detail program, silakan klik link di bawah ini ya\n${activityPageLink}\n\nTerima kasih atas perhatiannya.\nSalam,\nPT. Sempurna`
    else return `Selamat datang di ${programTitle}!\n\n${ongoingWording} pada tanggal ${date} pukul ${startTime} hingga ${endTime} WIB. ${conferenceLink}\n\nUntuk melihat detail program, silakan klik link di bawah ini ya\n${activityPageLink}\n\nTerima kasih atas perhatiannya.\nSalam,\nPT. Sempurna`;
  }

  private descriptionPartialConstructor(
    payload: GoogleCalendarDescriptionDto
  ): string {

    let description = payload.description
    if (description) {
      if (payload.programName) {
        const pattern = /Selamat datang di (.*)!\n/;
        description = description.replace(pattern, `Selamat datang di ${payload.programName}!\n`);
      }
      if (payload.startDateTime) {
        const date = moment(payload.startDateTime).format('YYYY-MM-DD')
        const startTime = moment(payload.startDateTime).format('HH:mm')
        const pattern = /tanggal (\d{4}-\d{2}-\d{2}) pukul (\d{2}:\d{2})/;
        description = description.replace(pattern, `tanggal ${date} pukul ${startTime}`);
      }
      if (payload.endDateTime) {
        const endTime = moment(payload.endDateTime).format('HH:mm')
        const pattern = /hingga (\d{2}:\d{2})/;
        description = description.replace(pattern, `hingga ${endTime}`);
      }
      if (payload.isYoutubeLiveIdExist) {
        const pattern1 = /berlangsung (.*) pada/;
        description = description.replace(pattern1, `berlangsung pada`);
      } else {
        if (payload.conferenceLink) {
          const pattern1 = /berlangsung (.*) pada/;
          const pattern2 = /berlangsung pada/;
          const pattern3 = /WIB. (.*)\n/;
          description = description.replace(pattern1, `berlangsung melalui Zoom pada`);
          description = description.replace(pattern2, `berlangsung melalui Zoom pada`);
          description = description.replace(pattern3, `WIB. ${payload.conferenceLink}\n`);
        }
      }
      if (payload.location) {
        const pattern1 = /berlangsung (.*) pada/;
        const pattern2 = /WIB. (.*)\n/;
        description = description.replace(pattern1, `berlangsung secara luring pada`);
        description = description.replace(pattern2, `WIB. Tempat: ${payload.location}\n`);
      }
      if (payload.training || payload.program) {
        const pattern = /silakan klik link di bawah ini ya\n.+?\n\nTerima kasih atas perhatiannya\./
        const encrypted = this.encryptTrainingOrProgram(payload.training, payload.program)
        const activityPageLink = `${this.webHost}/kelasku?td=${encrypted}`
        description = description.replace(pattern, `silakan klik link di bawah ini ya\n${activityPageLink}\n\nTerima kasih atas perhatiannya.`)
      }
      return description
    }
  }

  private encryptTrainingOrProgram(training?: Training, program?: Program): string {
    const data = {
      programName: training ? training.title : program.name,
      productType: training ? training.productType : program.productType,
      programType: training ? training.deliveryType : this.defineProgramType(program)
    }
    const encryptedData = CryptUtil.encryptData(JSON.stringify(data), this.encryptKey)
    return encryptedData
  }

  private defineProgramType(program: Program): string {
    if (program.type == ProgramType.STRUCTURED) {
      return "Bimbingan"
    } else if (program.type == ProgramType.UNSTRUCTURED) {
      return "Konsultasi"
    } else if (program.type == ProgramType.ENTREPRENEURSHIP) {
      if (program.format == ProgramFormat.PROGRAM) {
        return "Pendampingan"
      } else {
        return "Alat Usaha"
      }
    } else if (program.type == ProgramType.SELFLEARNING) {
      return "Mandiri"
    }
  }

  private serializeCacheValue(data: any): string {
    const value = JSON.stringify(data)
    return value
  }

  private deserializeCacheValue(cacheValue: string): any {
    const value = JSON.parse(cacheValue)
    return value
  }

  // the limit of 2000 email of external domain will be lifted in 24h
  private isBlockedTimeFinished(blockedDate: Date): boolean {
    // blockedDate = new Date(blockedDate)
    const dateNow = new Date();

    // Calculate the difference in milliseconds
    const timeDifference = dateNow.getTime() - blockedDate.getTime();

    // Calculate the number of milliseconds in 24 hours
    const twentyFourHoursInMilliseconds = 24 * 60 * 60 * 1000;

    // Compare the time difference to 24 hours
    if (timeDifference > twentyFourHoursInMilliseconds) {
      // The difference is greater than 24 hours
      return true;
    } else {
      // The difference is not greater than 24 hours
      return false;
    }

  }
}