import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { ProcessorService } from './processor.service';

async function bootstrap() {
  // 1) Crear app HTTP (Express por defecto)
  const app = await NestFactory.create(AppModule, {
    logger: ['log', 'error', 'warn'],
  });

  // 2) Habilitar CORS (Vite en 5173)
  app.enableCors({
    origin: ['http://localhost:5173'],
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS',
    credentials: false,
  });

  // 3) Ejecutar el processor ANTES de levantar el servidor (bloqueante)
  try {
    const svc = app.get(ProcessorService);
    console.log('⚙️  Generando Excel (ProcessorService.run() antes del HTTP)...');
    await svc.run(); // <- bloquea hasta que termine
    console.log('✔ ProcessorService.run() finalizado (archivo generado)');
  } catch (e) {
    console.error('✖ Error ejecutando ProcessorService.run():', e);
    // Si quieres que el server NO inicie si falla el Excel, descomenta:
    // process.exit(1);
  }

  // 4) Levantar servidor HTTP
  const port = Number(process.env.PORT ?? 3000);
  await app.listen(port);
  const url = await app.getUrl();
  console.log(`✅ Backend escuchando en ${url}`);
}

bootstrap().catch((e) => {
  console.error('✖ Error al iniciar la aplicación:', e);
  process.exit(1);
});
