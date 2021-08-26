import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { ApiService } from './api.service';
import { ApiController } from './api.controller';

import { Sheet, SheetSchema } from '../schemas/sheet.schema';
import { Measure, MeasureSchema } from '../schemas/measure.schema';
import { Artefact, ArtefactSchema } from '../schemas/artefact.schema';
import { Budget, BudgetSchema } from '../schemas/budget.schema';

@Module({
  imports: [
    MongooseModule.forFeature([
      { name: 'Sheet', schema: SheetSchema },
      { name: 'Measure', schema: MeasureSchema },
      { name: 'Artefact', schema: ArtefactSchema },
      { name: 'Budget', schema: BudgetSchema },
    ]),
  ],
  controllers: [ApiController],
  providers: [ApiService],
})
export class ApiModule {}
