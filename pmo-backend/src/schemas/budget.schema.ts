import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { Document } from 'mongoose';

export type BudgetDocument = Budget & Document;

@Schema()
export class Budget {
  @Prop()
  monthlySpendings: [number];

  @Prop()
  approvedBudgetPerMonth: number;

  @Prop()
  year: number;
}

export const BudgetSchema = SchemaFactory.createForClass(Budget);