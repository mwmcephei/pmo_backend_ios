import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import * as mongoose from 'mongoose';


export type BudgetDetailDocument = BudgetDetail & mongoose.Document;

@Schema()
export class BudgetDetail {

  @Prop()
  totalApprovedBudget: number;
  @Prop()
  spentBudget: number;
  @Prop()
  invoicedBudget: number;
  @Prop()
  forecastBudge: number;
  @Prop()
  contractBudget: number;

}

export const BudgetDetailSchema = SchemaFactory.createForClass(BudgetDetail);


