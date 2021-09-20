import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import * as mongoose from 'mongoose';

export type KPIDocument = KPI & mongoose.Document;

@Schema()
export class KPI {
  @Prop()
  title: string;
  @Prop()
  target: number;
  @Prop()
  actuals: number;
  @Prop()
  baseline: number;
  @Prop()
  plan1: number;
  @Prop()
  plan2: number;
  @Prop()
  plan3: number;
  @Prop()
  plan4: number;
}

export const KPISchema = SchemaFactory.createForClass(KPI);
