import { Component, inject, OnInit } from '@angular/core';
import { DocumentGeneratorService } from '../../services/document-generator-service';
import Swal from 'sweetalert2';
@Component({
  selector: 'app-dashboard-component',
  imports: [],
  providers: [],
  templateUrl: './dashboard-component.html',
  styleUrl: './dashboard-component.css',
})
export class DashboardComponent implements OnInit {
  private generatorService = inject(DocumentGeneratorService);

  // Estado de los archivos
  excelFile: File | null = null;
  templateFile: File | null = null;

  // Estado de la UI
  isLoading = false;
  isDraggingExcel = false;
  isDraggingTemplate = false;

  ngOnInit() {
    Swal.fire({
      toast: true,
      position: 'top-end',
      icon: 'success',
      title: 'Sesión iniciada correctamente',
      showConfirmButton: false,
      timer: 3000,
      timerProgressBar: true,
    });
  }

  // --- Manejadores de Drag & Drop Nativos ---

  onDragOver(event: DragEvent, type: 'excel' | 'template') {
    event.preventDefault(); // Necesario para permitir el drop
    event.stopPropagation();
    if (type === 'excel') this.isDraggingExcel = true;
    if (type === 'template') this.isDraggingTemplate = true;
  }

  onDragLeave(event: DragEvent, type: 'excel' | 'template') {
    event.preventDefault();
    event.stopPropagation();
    if (type === 'excel') this.isDraggingExcel = false;
    if (type === 'template') this.isDraggingTemplate = false;
  }

  onDrop(event: DragEvent, type: 'excel' | 'template') {
    event.preventDefault();
    event.stopPropagation();

    if (type === 'excel') this.isDraggingExcel = false;
    if (type === 'template') this.isDraggingTemplate = false;

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.validateAndSetFile(files[0], type);
    }
  }

  onFileSelected(event: Event, type: 'excel' | 'template') {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.validateAndSetFile(input.files[0], type);
    }
  }

  private validateAndSetFile(file: File, type: 'excel' | 'template') {
    const isExcel = type === 'excel' && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'));
    const isWord =
      type === 'template' && (file.name.endsWith('.docx') || file.name.endsWith('.doc'));

    if (isExcel) {
      this.excelFile = file;
    } else if (isWord) {
      this.templateFile = file;
    } else {
      Swal.fire({
        icon: 'warning',
        title: 'Formato incorrecto',
        text: `Por favor, sube un archivo ${type === 'excel' ? '.xlsx' : '.docx'} válido.`,
      });
    }
  }

  removeFile(type: 'excel' | 'template') {
    if (type === 'excel') this.excelFile = null;
    if (type === 'template') this.templateFile = null;
  }

  // --- Lógica Principal ---

  generateDocuments() {
    if (!this.excelFile || !this.templateFile) return;

    this.isLoading = true;

    this.generatorService.generateDocuments(this.excelFile, this.templateFile).subscribe({
      next: (blob: Blob) => {
        this.downloadFile(blob, 'procesados.zip');
        Swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'success',
          title: 'Documentos generados correctamente',
          showConfirmButton: false,
          timer: 3000,
          timerProgressBar: true,
        });
        this.isLoading = false;
      },
      error: (err) => {
        console.error('Error al generar:', err);
        Swal.fire({
          icon: 'error',
          title: 'Error de procesamiento',
          text: 'Hubo un problema al procesar los archivos. Verifique que el Excel y la plantilla sean correctos.',
        });
        this.isLoading = false;
      },
    });
  }

  private downloadFile(blob: Blob, filename: string) {
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    window.URL.revokeObjectURL(url);
  }
}
