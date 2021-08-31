import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import * as mongoose from 'mongoose';
import { Artefact } from './artefact.schema';

export type MeasureDocument = Measure & mongoose.Document;

@Schema()
export class Measure {
  @Prop()
  title: string;
  @Prop()
  id: number;
  @Prop()
  name: string;

  @Prop()
  focusArea: string;

  @Prop()
  measureLead: string;
  @Prop()
  measureSponsor: string;
  @Prop()
  lineOrgSponsor: string;
  @Prop()
  solutionManager: string;

  @Prop()
  approved: number;
  @Prop()
  spent: number;

  @Prop()
  kpiName: string;

  @Prop()
  actuals: number;
  @Prop()
  target: number;

  @Prop()
  risk: number;
  @Prop()
  budget: number;
  @Prop()
  artefact: number;

  @Prop()
  kpiProgress: number;

  @Prop({ type: [mongoose.Schema.Types.ObjectId], ref: 'Artefact' })
  artefacts: [Artefact];
}

export const MeasureSchema = SchemaFactory.createForClass(Measure);
