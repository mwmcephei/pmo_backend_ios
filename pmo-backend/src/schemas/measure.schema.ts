import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import * as mongoose from 'mongoose';
import { Artefact } from './artefact.schema';

export type MeasureDocument = Measure & mongoose.Document;

@Schema()
export class Measure {
  @Prop()
  title: string;

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