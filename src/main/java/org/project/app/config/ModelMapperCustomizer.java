package org.project.app.config;

import org.modelmapper.ModelMapper;

public interface ModelMapperCustomizer {
    void customize(ModelMapper mapper);
}
