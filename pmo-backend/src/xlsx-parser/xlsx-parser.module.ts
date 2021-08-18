import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { XlsxParserService } from './xlsx-parser.service';
import { XlsxParserController } from './xlsx-parser.controller';
import { Sheet, SheetSchema } from '../schemas/sheet.schema';
import { Measure, MeasureSchema } from '../schemas/measure.schema';
import { Artefact, ArtefactSchema } from '../schemas/artefact.schema';
import { Budget, BudgetSchema } from '../schemas/budget.schema';

@Module({
  imports: [
    MongooseModule.forFeature([
      { name: "Sheet", schema: SheetSchema },
      { name: "Measure", schema: MeasureSchema },
      { name: "Artefact", schema: ArtefactSchema },
      { name: "Budget", schema: BudgetSchema },
    ])
  ],
  providers: [XlsxParserController, XlsxParserService],
  controllers: [XlsxParserController]
})
export class XlsxParserModule {}
