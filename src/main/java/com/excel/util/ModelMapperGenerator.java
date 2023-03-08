package com.excel.util;

import com.excel.util.mapper.DoubleToDateConverter;
import com.excel.util.mapper.DoubleToStringConverter;
import com.excel.util.mapper.StringToDateConverter;
import com.excel.util.mapper.StringToStringConverter;
import org.modelmapper.ModelMapper;
import org.modelmapper.config.Configuration;
import org.modelmapper.convention.MatchingStrategies;

/**
 * The type Model mapper generator.
 */
public class ModelMapperGenerator {

    /**
     * Default model mapper model mapper.
     *
     * @return the model mapper
     */
    public static ModelMapper defaultModelMapper() {

        ModelMapper mapper = new ModelMapper();
        mapper.addConverter(new DoubleToDateConverter());
        mapper.addConverter(new StringToDateConverter());
        mapper.addConverter(new StringToStringConverter());
        mapper.addConverter(new DoubleToStringConverter());
        return mapper;
    }

    /**
     * Strict model mapper model mapper.
     *
     * @return the model mapper
     */
    public static ModelMapper strictModelMapper() {
        ModelMapper modelMapper = defaultModelMapper();
        modelMapper.getConfiguration().setMatchingStrategy(MatchingStrategies.STRICT);
        return modelMapper;
    }

    /**
     * Enable field model mapper model mapper.
     *
     * @param accessLevel the access level
     * @return the model mapper
     */
    public static ModelMapper enableFieldModelMapper(Configuration.AccessLevel accessLevel) {
        ModelMapper modelMapper = strictModelMapper();
        modelMapper.getConfiguration().setFieldMatchingEnabled(true);
        modelMapper.getConfiguration().setFieldAccessLevel(accessLevel);
        return modelMapper;
    }

    /**
     * Enable field model mapper model mapper.
     *
     * @return the model mapper
     */
    public static ModelMapper enableFieldModelMapper() {
        return enableFieldModelMapper(Configuration.AccessLevel.PRIVATE);
    }
}
