import { Controller, Get, Post,  Body, Req, Res } from '@nestjs/common';
import { AppService } from './app.service';
//import PushNotifications from 'node-pushnotifications';
import {Request, Response} from 'express';


@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

/*

  @Get()
  getHello(@Req() req: Request): string {
    return this.appService.getHello();


  }
*/
/*
  @Post("subscribe")
  getHello(@Req() req: Request): string {
    return "subscribe"
    
    // put into env
    const publicVapidKey = "BBzJAQ1MjGkiIxdBEQQX5KiR8lyFOwyZTyuYlcqMmbnAAEWT6VHAld3Vvenm_LBcz3QdHhdrM6Q-nV-OSht_MuY"
    const privateVapidKey = "JYjius3QYA5BRyrG6caZQRp8lXJe5V5OmhUAo71rjE4"

    const webpush = require('web-push');
 //   webpush.setGCMAPIKey('<Your GCM API Key Here>');      // ???? android

      webpush.setVapidDetails(
        'mailto:example@yourdomain.org',
        publicVapidKey,
        privateVapidKey
      );
      const declaration = req
      console.log(req.body)

      const payload = JSON.stringify({title: "moinsen"})
  //    webpush.sendNotification(declaration, payload).catch(err => console.log(err))

 //   return this.appService.getHello();

  }
 */


}
