package org.project.app.user.controller;

import lombok.RequiredArgsConstructor;
import org.project.app.user.service.UserMediaService;
import org.project.app.user.enums.UserMediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

@RequestMapping("/api/users/media")
@RequiredArgsConstructor
@RestController
public class UserMediaController {

    private final UserMediaService userMediaService;

    @GetMapping("/{id}")
    public ResponseEntity<Map<UserMediaType, String>> getAllMediaUrls(@PathVariable String id) {
        return ResponseEntity.ok(userMediaService.getAllMediaUrls(id));
    }

    @GetMapping("/{id}/file/{type}")
    public ResponseEntity<String> getUserMediaUrl(@PathVariable String id, @PathVariable UserMediaType type) {
        return ResponseEntity.ok(userMediaService.getMediaUrl(id, type));
    }

    @PostMapping("/{id}/file/{type}")
    public ResponseEntity<String> uploadUserMedia(@PathVariable String id, @PathVariable UserMediaType type, @RequestParam MultipartFile file) {
        userMediaService.uploadUserMedia(id, file, type);
        return ResponseEntity.ok("Media uploaded");
    }

}
