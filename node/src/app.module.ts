import { Module } from '@nestjs/common';
import { ProcessorService } from './processor.service';
import { IndicadoresController } from './indicadores/indicadores.controller';

@Module({
  imports: [],
  controllers: [IndicadoresController], // 👈 registra endpoints
  providers: [ProcessorService],        // 👈 tu servicio que genera el Excel
})
export class AppModule {}

