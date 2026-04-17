/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import java.io.IOException;
import java.util.NoSuchElementException;
import org.geotools.api.data.FeatureReader;
import org.geotools.api.data.Query;
import org.geotools.api.feature.simple.SimpleFeature;
import org.geotools.api.feature.simple.SimpleFeatureType;
import org.geotools.data.store.ContentEntry;

public class ExcelFeatureReader implements FeatureReader<SimpleFeatureType, SimpleFeature> {

    private final SimpleFeatureType featureType;

    public ExcelFeatureReader(ContentEntry entry, Query query) throws IOException {
        this.featureType = ((ExcelDataStore) entry.getDataStore()).getSchema();
    }

    @Override
    public SimpleFeatureType getFeatureType() {
        return this.featureType;
    }

    @Override
    public SimpleFeature next() throws IOException, IllegalArgumentException, NoSuchElementException {
        throw new NoSuchElementException("ExcelDataStore is write-only and does not contain readable features");
    }

    @Override
    public boolean hasNext() {
        return false;
    }

    @Override
    public void close() throws IOException {
        // no-op: this reader does not hold resources
    }
}
