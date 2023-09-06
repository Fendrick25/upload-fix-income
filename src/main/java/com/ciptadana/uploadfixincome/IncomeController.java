package com.ciptadana.uploadfixincome;

import lombok.RequiredArgsConstructor;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("api/cmf")
@RequiredArgsConstructor
@CrossOrigin
public class IncomeController {
    private final IncomeService incomeService;

    @PostMapping(value = "/uploads/fix-income", produces = MediaType.IMAGE_JPEG_VALUE)
    public ResponseEntity<byte[]> uploadFixIncome(@RequestParam("file") MultipartFile file) throws IOException {
        BufferedImage image = incomeService.processExcel(file);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "jpeg", baos);
        byte[] bytes = baos.toByteArray();

        return ResponseEntity
                .ok()
                .contentType(MediaType.IMAGE_JPEG)
                .body(bytes);
    }
}
