import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { XlsxParserModule } from './xlsx-parser/xlsx-parser.module';
import { MongooseModule } from '@nestjs/mongoose';

@Module({
  imports: [
    MongooseModule.forRoot('mongodb+srv://mwm:matthias88@cluster0.f8xt0.mongodb.net/myFirstDatabase?retryWrites=true&w=majority'),
    XlsxParserModule,
    //   ServeStaticModule.forRoot({
    //      rootPath: join(__dirname, '', 'client'),
    //    rootPath: join("/Users/mwm/Desktop/PMO/pmo/packages/pmo-backend/src/", '', 'client'),  
    //  }),


  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule { }
