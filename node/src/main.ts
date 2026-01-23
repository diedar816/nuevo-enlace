
import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { ProcessorService } from './processor.service';

async function bootstrap() {
  const app = await NestFactory.createApplicationContext(AppModule, {
    logger: ['log', 'error', 'warn'],
  });
  try {
    const svc = app.get(ProcessorService);
    await svc.run(); // Ejecuta el procesamiento directo por CLI
  } catch (e) {
    console.error(e);
  } finally {
    await app.close();
  }
}
bootstrap();
