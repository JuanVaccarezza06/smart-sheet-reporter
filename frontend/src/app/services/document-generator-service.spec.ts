import { TestBed } from '@angular/core/testing';

import { DocumentGeneratorService } from './document-generator-service';

describe('DocumentGeneratorService', () => {
  let service: DocumentGeneratorService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(DocumentGeneratorService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
