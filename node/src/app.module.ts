
import { Module } from '@nestjs/common';
import { ProcessorService } from './processor.service';

@Module({
  providers: [ProcessorService],
})
export class AppModule {}
