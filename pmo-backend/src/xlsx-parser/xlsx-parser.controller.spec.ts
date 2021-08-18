import { Test, TestingModule } from '@nestjs/testing';
import { XlsxParserController } from './xlsx-parser.controller';

describe('XlsxParserController', () => {
  let controller: XlsxParserController;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      controllers: [XlsxParserController],
    }).compile();

    controller = module.get<XlsxParserController>(XlsxParserController);
  });

  it('should be defined', () => {
    expect(controller).toBeDefined();
  });
});
