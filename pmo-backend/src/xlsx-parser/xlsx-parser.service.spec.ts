import { Test, TestingModule } from '@nestjs/testing';
import { XlsxParserService } from './xlsx-parser.service';

describe('XlsxParserService', () => {
  let service: XlsxParserService;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [XlsxParserService],
    }).compile();

    service = module.get<XlsxParserService>(XlsxParserService);
  });

  it('should be defined', () => {
    expect(service).toBeDefined();
  });
});
